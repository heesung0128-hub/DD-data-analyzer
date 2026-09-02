// --- 모델 타입 한글 라벨 ---
    const MODEL_LABELS = {
        linear: '선형(1차)', exponential: '지수', logarithmic: '로그',
        poly2: '2차 다항', poly3: '3차 다항', poly4: '4차 다항'
    };

    // --- 화면 라우팅 (메뉴 전환) ---
    function switchView(viewName) {
        document.getElementById('view-step1').classList.add('hidden');
        document.getElementById('view-step2').classList.add('hidden');
        document.getElementById('view-step3').classList.add('hidden');
        document.getElementById('view-step4').classList.add('hidden');
        document.getElementById('nav-step1').classList.remove('active');
        document.getElementById('nav-step2').classList.remove('active');
        document.getElementById('nav-step3').classList.remove('active');
        document.getElementById('nav-step4').classList.remove('active');

        document.getElementById('view-' + viewName).classList.remove('hidden');
        document.getElementById('nav-' + viewName).classList.add('active');

        if(viewName === 'step3' && window.currentChartData) {
            document.getElementById('predict-input-label').textContent = `예측하고 싶은 [${window.currentChartData.xLabel}] 값을 입력하세요:`;
            renderPredictEquation();
        }
        if(viewName === 'step4' && window.currentChartData) {
            renderIntegralEquationPreview();
        }
        refreshChartsForView(viewName);
    }

    function proceedToStep2() {
        document.getElementById('nav-step2').style.pointerEvents = 'auto';
        document.getElementById('nav-step2').style.opacity = '1';
        switchView('step2');
    }

    function proceedToStep3() {
        document.getElementById('nav-step3').style.pointerEvents = 'auto';
        document.getElementById('nav-step3').style.opacity = '1';
        switchView('step3');
    }

    // --- 차트 렌더링이 필요한 뷰로 들어갈 때/테마 전환 시 해당 뷰의 차트 다시 그리기 ---
    function refreshChartsForView(viewName) {
        if (!window.currentChartData) return;
        if (viewName === 'step2') {
            renderMainChart();
            if (window.currentChartData.residuals) {
                renderResidualChart(window.currentChartData.xData, window.currentChartData.residuals, window.currentChartData.xLabel);
            }
        } else if (viewName === 'step3') {
            renderPredictChart();
        } else if (viewName === 'step4') {
            if (window.currentIntegralRange) {
                renderIntegralChart(window.currentIntegralRange.lo, window.currentIntegralRange.hi);
            } else {
                renderStep4PreviewChart();
            }
        }
    }

    // --- 테마 토글 ---
    function toggleTheme() {
        const current = document.documentElement.getAttribute('data-theme');
        const icon = document.getElementById('theme-icon');
        if (current === 'dark') {
            document.documentElement.removeAttribute('data-theme');
            icon.classList.replace('fa-sun', 'fa-moon');
        } else {
            document.documentElement.setAttribute('data-theme', 'dark');
            icon.classList.replace('fa-moon', 'fa-sun');
        }

        const activeNav = document.querySelector('.nav-item.active');
        if (activeNav) refreshChartsForView(activeNav.id.replace('nav-', ''));
    }

    // --- 글로벌 변수 ---
    let globalData = [];
    let numericColumns = [];
    let currentFile = null;

    // --- 파일 업로드 및 인코딩 ---
    const fileInput = document.getElementById('fileInput');
    const dropzone = document.getElementById('dropzone');

    dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.style.transform = "scale(1.02)"; });
    dropzone.addEventListener('dragleave', (e) => { e.preventDefault(); dropzone.style.transform = "scale(1)"; });
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault(); dropzone.style.transform = "scale(1)";
        if (e.dataTransfer.files.length) {
            currentFile = e.dataTransfer.files[0];
            handleFile(currentFile);
        }
    });
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            currentFile = e.target.files[0];
            handleFile(currentFile);
        }
    });

    function reparseCurrentFile() {
        if (currentFile) handleFile(currentFile);
    }

    function updateDropzoneUI(labelHtml, descText) {
        const icon = document.getElementById('dropzone-icon');
        icon.className = 'fa-solid fa-file-circle-check';
        icon.style.color = 'var(--success)';
        document.getElementById('dropzone-title').innerHTML = labelHtml;
        document.getElementById('dropzone-desc').textContent = descText || '다른 파일을 분석하려면 다시 클릭하거나 드래그하세요.';
    }

    function handleFile(file) {
        if (!file) return;
        updateDropzoneUI(`선택된 파일: <span style="color: var(--primary);">${file.name}</span>`);

        const ext = file.name.split('.').pop().toLowerCase();

        if (ext === 'xlsx' || ext === 'xls') {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[firstSheetName];
                    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                    finalizeParsedData(json);
                } catch (err) {
                    alert("엑셀 파일을 읽는 중 오류가 발생했습니다. 파일 형식을 확인해 주세요.");
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            const encoding = document.getElementById('encoding-select').value;
            Papa.parse(file, {
                header: true,
                dynamicTyping: false, // 타입 자동 변환 해제로 파싱 속도 대폭 향상
                worker: true,         // 웹 워커를 사용하여 브라우저 멈춤 방지
                skipEmptyLines: true,
                encoding: encoding,
                complete: function(results) {
                    finalizeParsedData(results.data);
                },
                error: function(err) {
                    alert("파일을 읽는 중 오류가 발생했습니다. 인코딩 설정을 확인해 보세요.");
                }
            });
        }
    }

    // --- 클립보드 붙여넣기 ---
    function togglePasteArea() {
        document.getElementById('paste-area-wrapper').classList.toggle('hidden');
    }

    function handlePastedData() {
        const text = document.getElementById('pasteArea').value.trim();
        if (!text) {
            alert("붙여넣은 데이터가 없습니다. 엑셀/구글시트에서 표를 복사한 뒤 붙여넣어 주세요.");
            return;
        }
        Papa.parse(text, {
            header: true,
            dynamicTyping: false,
            skipEmptyLines: true,
            complete: function(results) {
                if (!results.data.length || Object.keys(results.data[0]).length < 2) {
                    alert("표 형식을 인식하지 못했습니다. 헤더를 포함해 2개 이상의 열을 붙여넣어 주세요.");
                    return;
                }
                currentFile = null;
                updateDropzoneUI('붙여넣기로 불러온 데이터', '다른 데이터를 사용하려면 파일을 올리거나 다시 붙여넣으세요.');
                finalizeParsedData(results.data);
            },
            error: function() {
                alert("붙여넣은 데이터를 해석하지 못했습니다. 표 형식을 확인해 주세요.");
            }
        });
    }

    // --- 샘플 데이터셋 ---
    function loadSampleDataset(name) {
        let raw = [];
        let cols = [];
        let label = '';

        if (name === 'sleep') {
            label = '수면시간 vs 학업성적 샘플 (선형 관계 연습용)';
            cols = ['수면시간', '학업성적'];
            raw = [[7.5,85],[6,72],[5.5,65],[8,90],[4.5,50],[7,80],[6.5,78],[5,60],[8.5,92],[6,70],
                   [7,82],[5.5,68],[6.5,75],[7.5,88],[5,62],[6,74],[8,86],[4,45],[7,81],[6.5,76],
                   [9,95],[5.8,71],[6.8,80],[7.2,84],[5.2,61],[6.2,73],[8.2,89],[4.8,58],[7.8,87],[6.4,72]];
        } else if (name === 'temp') {
            label = '기온 vs 음료 판매량 샘플 (2차 함수 관계 연습용)';
            cols = ['기온(℃)', '음료_판매량(개)'];
            raw = [[2,20],[5,35],[8,55],[11,70],[14,95],[17,130],[20,170],[23,205],[26,240],[29,260],
                   [32,255],[35,230],[10,60],[15,105],[19,155],[22,195],[27,250],[30,262],[6,40],[13,88],
                   [18,140],[24,220],[28,258],[31,258],[9,58],[16,120],[21,180],[25,235],[12,80],[33,240]];
        } else if (name === 'height') {
            label = '키 vs 발 크기 샘플 (선형 관계 연습용)';
            cols = ['키(cm)', '발_크기(mm)'];
            raw = [[150,225],[155,230],[158,232],[160,235],[162,238],[165,240],[168,243],[170,245],[172,248],[175,250],
                   [178,253],[180,255],[152,227],[157,231],[163,238],[167,242],[171,246],[174,249],[177,252],[159,234],
                   [161,236],[164,239],[169,244],[173,248],[176,251],[179,254],[156,230],[166,241],[181,256],[153,228]];
        } else if (name === 'bacteria') {
            label = '경과시간 vs 세균 수 샘플 (지수 함수 관계 연습용)';
            cols = ['경과시간(시간)', '세균_수(마리)'];
            raw = [[0,50],[1,68],[2,90],[3,123],[4,166],[5,225],[6,300],[7,410],[8,545],
                   [9,735],[10,990],[11,1330],[12,1790],[13,2410],[14,3250]];
        } else {
            return;
        }

        const data = raw.map(r => ({ [cols[0]]: r[0], [cols[1]]: r[1] }));
        currentFile = null;
        updateDropzoneUI(`선택된 샘플: <span style="color: var(--primary);">${label}</span>`);
        finalizeParsedData(data);
    }

    // --- 파싱 결과 공통 처리 (CSV/XLSX/붙여넣기/샘플 공용) ---
    function finalizeParsedData(data) {
        globalData = data;
        renderTable(globalData);
        processColumns(globalData);
    }

    // --- 데이터 테이블 렌더링 ---
    function renderTable(data) {
        if(data.length === 0) return;
        const container = document.getElementById('table-container');
        let html = '<table><thead><tr>';
        const headers = Object.keys(data[0]);
        headers.forEach(h => html += `<th>${h}</th>`);
        html += '</tr></thead><tbody>';

        const rowsToShow = Math.min(data.length, 5);
        for(let i=0; i<rowsToShow; i++) {
            html += '<tr>';
            headers.forEach(h => {
                html += `<td>${data[i][h] !== null ? data[i][h] : ''}</td>`;
            });
            html += '</tr>';
        }
        html += '</tbody></table>';
        container.innerHTML = html;
        document.getElementById('preview-card').classList.remove('hidden');
    }

    // --- 변수 필터링 ---
    function processColumns(data) {
        if(data.length === 0) return;
        const headers = Object.keys(data[0]);
        numericColumns = headers.filter(h => {
            for(let i=0; i<Math.min(5, data.length); i++) {
                let val = data[i][h];
                if (typeof val === 'number') return true;
                if (typeof val === 'string' && !isNaN(parseFloat(val.replace(/,/g, '')))) return true;
            }
            return false;
        });

        const colError = document.getElementById('col-error');
        const colSelectors = document.getElementById('col-selectors');
        const analyzeAction = document.getElementById('analyze-action');

        if(numericColumns.length < 2) {
            colError.classList.remove('hidden');
            colSelectors.classList.add('hidden');
            analyzeAction.classList.add('hidden');
            document.getElementById('nav-step2').style.pointerEvents = 'none';
            document.getElementById('nav-step2').style.opacity = '0.5';
        } else {
            colError.classList.add('hidden');
            colSelectors.classList.remove('hidden');
            analyzeAction.classList.remove('hidden');

            populateSelect('xCol', numericColumns, numericColumns[0]);
            populateSelect('yCol', numericColumns, numericColumns[1]);
        }
    }

    function populateSelect(id, options, defaultVal) {
        const select = document.getElementById(id);
        select.innerHTML = '';
        options.forEach(opt => {
            const el = document.createElement('option');
            el.value = opt; el.textContent = opt;
            if(opt === defaultVal) el.selected = true;
            select.appendChild(el);
        });
    }

    document.getElementById('xCol').addEventListener('change', (e) => {
        const xVal = e.target.value;
        const yVal = document.getElementById('yCol').value;
        populateSelect('yCol', numericColumns.filter(c => c !== xVal), yVal === xVal ? numericColumns.filter(c => c !== xVal)[0] : yVal);
    });

    // --- 일반화된 지수 함수 찾기 y = a * e^(bx) + c ---
    function findGeneralizedExponential(points) {
        const yVals = points.map(p => p[1]);
        const minY = Math.min(...yVals);
        const maxY = Math.max(...yVals);
        const range = Math.max(maxY - minY, 0.001);

        let bestSSR = Infinity;
        let bestModel = null;

        function evaluate(c, sign) {
            let transformed = [];
            for(let i=0; i<points.length; i++) {
                let y = points[i][1] - c;
                if(sign > 0 && y <= 0) return;
                if(sign < 0 && y >= 0) return;
                transformed.push([points[i][0], sign > 0 ? y : -y]);
            }

            let reg = regression.exponential(transformed, { precision: 10 });
            if(isNaN(reg.equation[0]) || isNaN(reg.equation[1])) return;

            let a = sign > 0 ? reg.equation[0] : -reg.equation[0];
            let b = reg.equation[1];

            let ssr = 0;
            for(let i=0; i<points.length; i++) {
                let pred = a * Math.exp(b * points[i][0]) + c;
                ssr += Math.pow(points[i][1] - pred, 2);
            }

            if (ssr < bestSSR) {
                bestSSR = ssr;
                bestModel = { a, b, c };
            }
        }

        const steps = 1000;
        for(let i = 0; i <= steps; i++) {
            let c = minY - 0.0001 - (range * 10 * i / steps);
            evaluate(c, 1);
        }
        for(let i = 0; i <= steps; i++) {
            let c = maxY + 0.0001 + (range * 10 * i / steps);
            evaluate(c, -1);
        }
        evaluate(0, 1);
        evaluate(0, -1);

        if (!bestModel) return null;

        let sst = 0;
        const meanY = yVals.reduce((sum, y) => sum + y, 0) / yVals.length;
        for(let i=0; i<yVals.length; i++) {
            sst += Math.pow(yVals[i] - meanY, 2);
        }

        let r2 = sst === 0 ? 1 : 1 - (bestSSR / sst);

        let aStr = bestModel.a.toPrecision(4);
        let bStr = bestModel.b.toPrecision(4);
        let cStr = Math.abs(bestModel.c).toPrecision(4);
        let eqStr = `${aStr} e^(${bStr}x) ${bestModel.c >= 0 ? '+' : '-'} ${cStr}`;
        if(Math.abs(bestModel.c) < 1e-10) eqStr = `${aStr} e^(${bStr}x)`;

        return {
            equation: [bestModel.a, bestModel.b, bestModel.c],
            string: eqStr,
            r2: r2,
            predict: function(x) {
                return [x, bestModel.a * Math.exp(bestModel.b * x) + bestModel.c];
            }
        };
    }

    // --- 일반화된 로그 함수 찾기 y = a + b * ln(x - c) ---
    function findGeneralizedLogarithmic(points) {
        const xVals = points.map(p => p[0]);
        const yVals = points.map(p => p[1]);
        const minX = Math.min(...xVals);
        const rangeX = Math.max(Math.max(...xVals) - minX, 0.001);

        let bestSSR = Infinity;
        let bestModel = null;

        function evaluate(c) {
            let transformed = [];
            for(let i=0; i<points.length; i++) {
                let x = points[i][0] - c;
                if(x <= 0) return;
                transformed.push([x, points[i][1]]);
            }

            let reg = regression.logarithmic(transformed, { precision: 10 });
            if(isNaN(reg.equation[0]) || isNaN(reg.equation[1])) return;

            let a = reg.equation[0];
            let b = reg.equation[1];

            let ssr = 0;
            for(let i=0; i<points.length; i++) {
                let pred = a + b * Math.log(points[i][0] - c);
                ssr += Math.pow(points[i][1] - pred, 2);
            }

            if (ssr < bestSSR) {
                bestSSR = ssr;
                bestModel = { a, b, c };
            }
        }

        const steps = 1000;
        for(let i = 0; i <= steps; i++) {
            let c = minX - 0.0001 - (rangeX * 10 * i / steps);
            evaluate(c);
        }
        evaluate(0);

        if (!bestModel) return null;

        let sst = 0;
        const meanY = yVals.reduce((sum, y) => sum + y, 0) / yVals.length;
        for(let i=0; i<yVals.length; i++) {
            sst += Math.pow(yVals[i] - meanY, 2);
        }

        let r2 = sst === 0 ? 1 : 1 - (bestSSR / sst);

        let aStr = bestModel.a.toPrecision(4);
        let bStr = Math.abs(bestModel.b).toPrecision(4);
        let cStr = Math.abs(bestModel.c).toPrecision(4);
        let signB = bestModel.b >= 0 ? '+' : '-';
        let signC = bestModel.c >= 0 ? '-' : '+';
        let eqStr = `${aStr} ${signB} ${bStr} ln(x ${signC} ${cStr})`;
        if(Math.abs(bestModel.c) < 1e-10) eqStr = `${aStr} ${signB} ${bStr} ln(x)`;

        return {
            equation: [bestModel.a, bestModel.b, bestModel.c],
            string: eqStr,
            r2: r2,
            predict: function(x) {
                return [x, bestModel.a + bestModel.b * Math.log(x - bestModel.c)];
            }
        };
    }

    // --- 모델 타입에 따라 회귀 결과 계산 (공용) ---
    function fitModel(modelType, points) {
        if(modelType === 'linear') {
            return regression.linear(points, { precision: 4 });
        } else if(modelType === 'exponential') {
            return findGeneralizedExponential(points) || regression.exponential(points, { precision: 4 });
        } else if(modelType === 'logarithmic') {
            return findGeneralizedLogarithmic(points) || regression.logarithmic(points, { precision: 4 });
        } else if(modelType === 'poly2') {
            return regression.polynomial(points, { order: 2, precision: 4 });
        } else if(modelType === 'poly3') {
            return regression.polynomial(points, { order: 3, precision: 4 });
        } else if(modelType === 'poly4') {
            return regression.polynomial(points, { order: 4, precision: 4 });
        }
        return null;
    }

    function isPolynomialFamily(modelType) {
        return modelType === 'linear' || modelType === 'poly2' || modelType === 'poly3' || modelType === 'poly4';
    }

    // --- 최적 모델 자동 추천 ---
    function runAutoRecommend() {
        const xCol = document.getElementById('xCol').value;
        const yCol = document.getElementById('yCol').value;
        if (!xCol || !yCol || xCol === yCol) {
            alert("독립변수와 종속변수를 서로 다르게 선택해주세요.");
            return;
        }

        let validRows = [];
        globalData.forEach(row => {
            let rawX = row[xCol];
            let rawY = row[yCol];
            let xVal = typeof rawX === 'string' ? parseFloat(rawX.replace(/,/g, '')) : rawX;
            let yVal = typeof rawY === 'string' ? parseFloat(rawY.replace(/,/g, '')) : rawY;
            if(typeof xVal === 'number' && !isNaN(xVal) && typeof yVal === 'number' && !isNaN(yVal)) {
                validRows.push({ x: xVal, y: yVal });
            }
        });

        if (validRows.length < 6) {
            alert("자동 추천을 실행하려면 최소 6개 이상의 데이터가 필요합니다.");
            return;
        }

        const useZeroAdjust = document.getElementById('zeroPointAdjust').checked;
        const minXOriginal = Math.min(...validRows.map(r => r.x));
        const shiftX = useZeroAdjust ? minXOriginal - 1 : 0;

        // 6개 모델을 공정하게 비교하기 위해, 모든 모델에 공통으로 적용 가능한 데이터(x-shift>0, y>0)만 사용
        const commonPoints = [];
        validRows.forEach(row => {
            const adjX = row.x - shiftX;
            if (adjX <= 0) return;
            if (row.y <= 0) return;
            commonPoints.push([adjX, row.y]);
        });
        const excludedCount = validRows.length - commonPoints.length;

        if (commonPoints.length < 6) {
            alert("모든 모델에 공통으로 쓸 수 있는 데이터(x>0, y>0 기준)가 부족해 자동 추천을 실행할 수 없습니다. 영점 조절 옵션을 켜보거나 데이터를 확인해주세요.");
            return;
        }

        const candidates = ['linear', 'exponential', 'logarithmic', 'poly2', 'poly3', 'poly4'];
        const n = commonPoints.length;
        let results = [];

        candidates.forEach(type => {
            try {
                const reg = fitModel(type, commonPoints);
                if (!reg || isNaN(reg.r2)) return;
                const p = reg.equation.length; // 사용된 파라미터(계수) 개수
                if (n - p - 1 <= 0) return; // 자유도 부족 시 비교 대상에서 제외
                const adjR2 = 1 - (1 - reg.r2) * (n - 1) / (n - p - 1);
                results.push({ type, r2: reg.r2, adjR2, p });
            } catch (e) { /* 해당 모델은 건너뜀 */ }
        });

        if (results.length === 0) {
            alert("자동 추천 계산에 실패했습니다. 데이터를 확인해주세요.");
            return;
        }

        results.sort((a, b) => b.adjR2 - a.adjR2);
        const best = results[0];

        let tableHtml = `<div class="alert alert-info" style="flex-direction:column;">
            <strong><i class="fa-solid fa-wand-magic-sparkles"></i> 자동 추천 결과: <span style="color:var(--primary);">${MODEL_LABELS[best.type]} 함수</span> 모델을 추천합니다</strong>
            <p class="text-muted mt-2" style="font-size:0.78rem; line-height:1.5;">
                공정한 비교를 위해 모든 모델에 공통으로 사용 가능한 데이터 ${n}개(x&gt;0, y&gt;0 기준${excludedCount > 0 ? `, 원본 대비 ${excludedCount}개 제외` : ''})를 사용했습니다.
                단순 R²가 아니라 모델의 항 개수(자유도)를 보정한 <b>조정된 R²(Adjusted R²)</b>로 비교했는데, 차수가 높을수록 R²는 항상 커지는 경향이 있어 그대로 비교하면 과적합(overfitting)된 모델이 부당하게 유리해지기 때문입니다.
            </p>
            <div class="table-container mt-2" style="max-height:none;">
            <table><thead><tr><th>모델</th><th>파라미터 수</th><th>R²</th><th>조정된 R²</th></tr></thead><tbody>`;
        results.forEach(r => {
            const isBest = r.type === best.type;
            tableHtml += `<tr style="${isBest ? 'font-weight:700; color: var(--primary);' : ''}">
                <td>${isBest ? '⭐ ' : ''}${MODEL_LABELS[r.type]}</td>
                <td>${r.p}</td>
                <td>${r.r2.toFixed(4)}</td>
                <td>${r.adjR2.toFixed(4)}</td>
            </tr>`;
        });
        tableHtml += `</tbody></table></div></div>`;

        const box = document.getElementById('auto-recommend-box');
        box.innerHTML = tableHtml;
        box.classList.remove('hidden');

        document.getElementById('trendModel').value = best.type;
    }

    // ==================== 수식 LaTeX 변환 (KaTeX) ====================

    function numToLatex(num) {
        if (num === undefined || num === null || isNaN(num)) return '0';
        let s = num.toPrecision(4);
        if (s.includes('e')) {
            const [mantissa, exp] = s.split('e');
            const expNum = parseInt(exp, 10);
            return `${mantissa} \\times 10^{${expNum}}`;
        }
        return s;
    }

    // 영점 조절이 켜져 있으면 x를 (x - shift) 형태로 표기
    function getXTermLatex(shiftX, useZeroAdjust) {
        if (useZeroAdjust && shiftX !== 0) {
            const sign = shiftX > 0 ? '-' : '+';
            return `\\left(x ${sign} ${numToLatex(Math.abs(shiftX))}\\right)`;
        }
        return 'x';
    }

    // 다항함수(선형~4차) 계수 배열 [최고차항, ..., 상수항] -> LaTeX
    function polynomialToLatex(equation, xTerm) {
        const order = equation.length - 1;
        const parts = [];
        for (let i = 0; i <= order; i++) {
            const power = order - i;
            const coef = equation[i];
            if (Math.abs(coef) < 1e-12) continue;
            const absStr = numToLatex(Math.abs(coef));
            let term;
            if (power === 0) term = absStr;
            else if (power === 1) term = `${absStr}${xTerm}`;
            else term = `${absStr}${xTerm}^{${power}}`;
            parts.push({ sign: coef >= 0 ? '+' : '-', term });
        }
        if (parts.length === 0) return 'y = 0';
        let s = (parts[0].sign === '-' ? '-' : '') + parts[0].term;
        for (let i = 1; i < parts.length; i++) s += ` ${parts[i].sign} ${parts[i].term}`;
        return `y = ${s}`;
    }

    // 다항함수 부정적분(원시함수) F(x) -> LaTeX
    function polynomialAntiderivativeLatex(equation, xTerm) {
        const order = equation.length - 1;
        const parts = [];
        for (let i = 0; i <= order; i++) {
            const power = order - i;
            const newPower = power + 1;
            const coef = equation[i] / newPower;
            if (Math.abs(coef) < 1e-12) continue;
            const absStr = numToLatex(Math.abs(coef));
            let term;
            if (newPower === 0) term = absStr;
            else if (newPower === 1) term = `${absStr}${xTerm}`;
            else term = `${absStr}${xTerm}^{${newPower}}`;
            parts.push({ sign: coef >= 0 ? '+' : '-', term });
        }
        if (parts.length === 0) return 'F(x) = 0 + C';
        let s = (parts[0].sign === '-' ? '-' : '') + parts[0].term;
        for (let i = 1; i < parts.length; i++) s += ` ${parts[i].sign} ${parts[i].term}`;
        return `F(x) = ${s} + C`;
    }

    // 모델 타입에 맞춰 현재 추세선 식을 LaTeX로 구성 (Step2/3/4 공용)
    function buildEquationLatex(d) {
        const xTerm = getXTermLatex(d.shiftX, d.useZeroAdjust);
        const eq = d.regResult.equation;

        if (isPolynomialFamily(d.modelType)) {
            return polynomialToLatex(eq, xTerm);
        } else if (d.modelType === 'exponential') {
            const [a, b, c] = eq;
            let s = `y = ${numToLatex(a)} e^{${numToLatex(b)}${xTerm}}`;
            if (c !== undefined && Math.abs(c) > 1e-9) {
                s += ` ${c >= 0 ? '+' : '-'} ${numToLatex(Math.abs(c))}`;
            }
            return s;
        } else if (d.modelType === 'logarithmic') {
            const [a, b, c] = eq;
            const inner = (c !== undefined && Math.abs(c) > 1e-9) ? `${xTerm} ${c >= 0 ? '-' : '+'} ${numToLatex(Math.abs(c))}` : xTerm;
            const signB = b >= 0 ? '+' : '-';
            return `y = ${numToLatex(a)} ${signB} ${numToLatex(Math.abs(b))} \\ln\\left(${inner}\\right)`;
        }
        return 'y = f(x)';
    }

    // KaTeX 렌더링 (실패 시 일반 텍스트로 대체)
    function renderLatex(el, latex, displayMode) {
        if (!el) return;
        try {
            katex.render(latex, el, { throwOnError: false, displayMode: !!displayMode });
        } catch (e) {
            el.textContent = latex;
        }
    }

    // ==================== 통계 분석 ====================

    function runAnalysis() {
        const xCol = document.getElementById('xCol').value;
        const yCol = document.getElementById('yCol').value;
        const modelType = document.getElementById('trendModel').value;
        const useZeroAdjust = document.getElementById('zeroPointAdjust').checked;

        // 1차 필터링: 유효한 숫자 데이터만 추출 (콤마 제거 포함)
        let validRows = [];
        globalData.forEach(row => {
            let rawX = row[xCol];
            let rawY = row[yCol];
            let xVal = typeof rawX === 'string' ? parseFloat(rawX.replace(/,/g, '')) : rawX;
            let yVal = typeof rawY === 'string' ? parseFloat(rawY.replace(/,/g, '')) : rawY;

            if(typeof xVal === 'number' && !isNaN(xVal) && typeof yVal === 'number' && !isNaN(yVal)) {
                validRows.push({ x: xVal, y: yVal });
            }
        });

        if(validRows.length < 3) {
            alert("데이터가 부족합니다.");
            return;
        }

        // 평행이동 처리 (영점 조절)
        const minXOriginal = Math.min(...validRows.map(r => r.x));
        let shiftX = 0;
        if (useZeroAdjust) {
            // 로그함수는 x가 무조건 양수여야 하므로 시작점을 1로 맞춥니다.
            shiftX = modelType === 'logarithmic' ? minXOriginal - 1 : minXOriginal;
        }

        let points = [];
        let xData = [];
        let yData = [];

        validRows.forEach(row => {
            let adjX = row.x - shiftX;
            let y = row.y;

            if(modelType === 'logarithmic' && adjX <= 0) return; // log(0) 에러 방지
            if(modelType === 'exponential' && y <= 0) return; // 지수함수는 y가 0 이하일 수 없음

            points.push([adjX, y]);
            xData.push(row.x);
            yData.push(y);
        });

        if(points.length < 3) {
            alert("선택한 수학 모델로 분석하기 위한 유효한 데이터가 부족합니다. (음수/0 값 제외됨)");
            return;
        }

        const order = modelType === 'poly2' ? 2 : modelType === 'poly3' ? 3 : modelType === 'poly4' ? 4 : null;
        if (order !== null && points.length < order + 3) {
            alert(`${MODEL_LABELS[modelType]} 함수는 최소 ${order + 3}개 이상의 데이터가 있어야 신뢰할 수 있는 분석이 가능합니다. (현재 ${points.length}개)`);
            return;
        }

        // Regression.js 추세선 및 결정계수(R2) 계산
        let regResult = fitModel(modelType, points);
        if(regResult.string.includes('y = ')) regResult.string = regResult.string.replace('y = ', '');
        // regression.js는 음수 계수를 "+ -3" 형태로 표기하므로 "- 3" 형태로 보기 좋게 정리 (다운로드 리포트용)
        regResult.string = regResult.string.replace(/\+ -/g, '- ');

        let equationText = regResult.string;

        // 다운로드 리포트용 평문 수식 (x 대신 평행이동한 (x - shift) 형태 명시)
        if (useZeroAdjust && shiftX !== 0) {
            let shiftStr = shiftX > 0 ? `- ${shiftX}` : `+ ${Math.abs(shiftX)}`;
            equationText = equationText.replace(/x/g, `(x ${shiftStr})`);
        }

        const r2 = regResult.r2;

        // jStat: p-value (선형 피어슨 기준, 참고용)
        const pearsonR = jStat.corrcoeff(xData, yData);
        const n = xData.length;
        const t_stat = pearsonR * Math.sqrt(n - 2) / Math.sqrt(1 - pearsonR * pearsonR);
        const p_value = 2 * (1 - jStat.studentt.cdf(Math.abs(t_stat), n - 2));

        // 곡선 그래프 그리기용 촘촘한 좌표 생성 (원래 X 좌표대로 그려지도록 복원)
        const minX = Math.min(...xData);
        const maxX = Math.max(...xData);
        let curveX = [];
        let curveY = [];
        const steps = 100;
        for(let i=0; i<=steps; i++) {
            let originalX = minX + (maxX - minX) * (i / steps);
            let adjX = originalX - shiftX;
            if(adjX <= 0 && modelType === 'logarithmic') adjX = 0.0001; // log(0) 방지

            let pred = regResult.predict(adjX);
            curveX.push(originalX); // 화면에는 원래 X축 좌표대로 그립니다.
            curveY.push(pred[1]);
        }

        // 잔차 계산 (실제값 - 예측값)
        let residuals = [];
        for(let i=0; i<xData.length; i++) {
            let adjX = xData[i] - shiftX;
            let pred = regResult.predict(adjX)[1];
            residuals.push(yData[i] - pred);
        }

        // UI 업데이트
        document.getElementById('result-card').classList.remove('hidden');
        renderLatex(document.getElementById('res-eq'), buildEquationLatex({ regResult, modelType, shiftX, useZeroAdjust }));
        document.getElementById('res-r2').textContent = r2.toFixed(4);
        document.getElementById('res-p').textContent = p_value.toFixed(4);

        const interpBox = document.getElementById('interpretation-box');
        let interpHtml = '';
        if (r2 >= 0.7) {
            interpHtml = `
                <div class="alert alert-success">
                    <i class="fa-solid fa-circle-check"></i>
                    <div>
                        <strong>강력한 설명력!</strong>
                        <p>결정계수(R²)가 ${r2.toFixed(2)}로 매우 높습니다. 선택하신 <b>[${MODEL_LABELS[modelType]}] 함수식</b>이 두 변수 간의 관계를 훌륭하게 설명하고 있습니다. 탐구 보고서에 이 함수식을 핵심 수학적 근거로 제시해 보세요!</p>
                    </div>
                </div>`;
        } else if (r2 >= 0.4) {
            interpHtml = `
                <div class="alert alert-info">
                    <i class="fa-solid fa-chart-bar"></i>
                    <div>
                        <strong>유의미한 설명력</strong>
                        <p>결정계수(R²)가 ${r2.toFixed(2)}로 중간 정도입니다. 선택하신 함수식이 어느 정도 의미 있는 패턴을 보여주지만, 예외적인 데이터(오차)도 제법 존재합니다. 다른 추세선 모델을 적용했을 때 더 높아지는지 비교해 보세요.</p>
                    </div>
                </div>`;
        } else {
            interpHtml = `
                <div class="alert alert-warning">
                    <i class="fa-solid fa-triangle-exclamation"></i>
                    <div>
                        <strong>설명력 부족</strong>
                        <p>결정계수(R²)가 ${r2.toFixed(2)}로 많이 낮습니다. 데이터가 선택한 함수식을 거의 따르지 않거나 다른 변수의 영향이 큽니다. 추세선 모델을 다른 함수(다항, 로그 등)로 변경해 보세요!</p>
                    </div>
                </div>`;
        }

        // 3차/4차 함수 + 데이터 수가 적을 때 과적합 경고
        if (order !== null && order >= 3 && points.length < (order + 1) * 3) {
            interpHtml += `
                <div class="alert alert-warning">
                    <i class="fa-solid fa-circle-exclamation"></i>
                    <div>
                        <strong>과적합(Overfitting) 주의</strong>
                        <p>데이터 개수(${points.length}개)에 비해 ${MODEL_LABELS[modelType]} 함수의 항이 많은 편입니다. 차수가 높을수록 곡선이 점을 억지로 통과해 R²가 높게 나올 수 있으므로, 잔차 그래프를 함께 확인하고 더 낮은 차수 모델과 비교해 보세요.</p>
                    </div>
                </div>`;
        }
        interpBox.innerHTML = interpHtml;

        // 차트 커스텀 옵션 초기화 (새 분석 시 이전 스타일 오버라이드 제거)
        window.currentChartOptions = null;
        document.getElementById('opt-title').value = '';
        document.getElementById('opt-title').placeholder = `y = ${equationText}`;
        document.getElementById('opt-xlabel').value = '';
        document.getElementById('opt-xlabel').placeholder = xCol;
        document.getElementById('opt-ylabel').value = '';
        document.getElementById('opt-ylabel').placeholder = yCol;
        document.getElementById('opt-xmin').value = '';
        document.getElementById('opt-xmax').value = '';
        document.getElementById('opt-ymin').value = '';
        document.getElementById('opt-ymax').value = '';

        window.currentChartData = {
            xData, yData, curveX, curveY,
            xLabel: xCol, yLabel: yCol,
            equation: equationText,
            regResult: regResult,
            shiftX: shiftX,
            useZeroAdjust: useZeroAdjust,
            modelType: modelType,
            residuals: residuals
        };
        window.currentIntegralRange = null;

        renderMainChart();
        renderResidualChart(xData, residuals, xCol);

        // 예측/적분 단계 잠금 해제
        document.getElementById('nav-step3').style.pointerEvents = 'auto';
        document.getElementById('nav-step3').style.opacity = '1';
        document.getElementById('nav-step4').style.pointerEvents = 'auto';
        document.getElementById('nav-step4').style.opacity = '1';
        document.getElementById('integral-result-box').classList.add('hidden');
        document.getElementById('integralA').value = '';
        document.getElementById('integralB').value = '';
    }

    // --- 차트 커스텀 옵션 ---
    function toggleChartOptions() {
        const body = document.getElementById('chart-options-body');
        const caret = document.getElementById('chart-options-caret');
        body.classList.toggle('hidden');
        caret.classList.toggle('fa-chevron-down');
        caret.classList.toggle('fa-chevron-up');
    }

    function applyChartOptions() {
        if (!window.currentChartData) return;
        const opts = {};
        const title = document.getElementById('opt-title').value.trim();
        const xlabel = document.getElementById('opt-xlabel').value.trim();
        const ylabel = document.getElementById('opt-ylabel').value.trim();
        const xmin = document.getElementById('opt-xmin').value;
        const xmax = document.getElementById('opt-xmax').value;
        const ymin = document.getElementById('opt-ymin').value;
        const ymax = document.getElementById('opt-ymax').value;

        if (title) opts.title = title;
        if (xlabel) opts.xLabel = xlabel;
        if (ylabel) opts.yLabel = ylabel;
        opts.pointColor = document.getElementById('opt-point-color').value;
        opts.lineColor = document.getElementById('opt-line-color').value;
        opts.lineWidth = parseFloat(document.getElementById('opt-line-width').value) || 3;
        if (xmin !== '' && xmax !== '') opts.xRange = [parseFloat(xmin), parseFloat(xmax)];
        if (ymin !== '' && ymax !== '') opts.yRange = [parseFloat(ymin), parseFloat(ymax)];

        window.currentChartOptions = opts;
        renderMainChart();
    }

    // --- 그래프 렌더링 (Step2/3/4 공용) ---
    function renderChart(containerId, xData, yData, curveX, curveY, xLabel, yLabel, overrides) {
        overrides = overrides || {};
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        const textColor = isDark ? '#e2e8f0' : '#1a202c';
        const gridColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)';
        const paperColor = 'rgba(0,0,0,0)';

        const scatter = {
            x: xData, y: yData, mode: 'markers', type: 'scatter', name: '실제 데이터',
            marker: { color: overrides.pointColor || 'hsl(255, 65%, 60%)', size: 8, opacity: 0.7 }
        };

        const trendline = {
            x: curveX, y: curveY, mode: 'lines', type: 'scatter', name: '예측 곡선',
            line: { color: overrides.lineColor || 'hsl(5, 75%, 55%)', dash: 'dash', width: overrides.lineWidth || 3 }
        };

        const xaxis = { title: overrides.xLabel || xLabel, gridcolor: gridColor, zerolinecolor: gridColor };
        const yaxis = { title: overrides.yLabel || yLabel, gridcolor: gridColor, zerolinecolor: gridColor };
        if (overrides.xRange) xaxis.range = overrides.xRange;
        if (overrides.yRange) yaxis.range = overrides.yRange;

        const layout = {
            autosize: true, plot_bgcolor: paperColor, paper_bgcolor: paperColor,
            font: { color: textColor, family: 'Inter' },
            xaxis: xaxis,
            yaxis: yaxis,
            margin: { l: 50, r: 20, t: overrides.title ? 40 : 15, b: 50 },
            legend: { orientation: "h", y: -0.2 }
        };
        if (overrides.title) layout.title = { text: overrides.title, font: { family: 'Outfit', size: 16 } };

        Plotly.newPlot(containerId, [scatter, trendline], layout, { responsive: true, displayModeBar: true, displaylogo: false });
    }

    function renderMainChart() {
        if (!window.currentChartData) return;
        const d = window.currentChartData;
        renderChart('chart-container', d.xData, d.yData, d.curveX, d.curveY, d.xLabel, d.yLabel, window.currentChartOptions);
    }

    function renderPredictChart() {
        if (!window.currentChartData) return;
        const d = window.currentChartData;
        renderChart('predict-chart-container', d.xData, d.yData, d.curveX, d.curveY, d.xLabel, d.yLabel, window.currentChartOptions);
    }

    function renderStep4PreviewChart() {
        if (!window.currentChartData) return;
        const d = window.currentChartData;
        renderChart('integral-chart-container', d.xData, d.yData, d.curveX, d.curveY, d.xLabel, d.yLabel, window.currentChartOptions);
    }

    function renderPredictEquation() {
        if (!window.currentChartData) return;
        renderLatex(document.getElementById('predict-eq-latex'), buildEquationLatex(window.currentChartData));
    }

    function renderIntegralEquationPreview() {
        if (!window.currentChartData) return;
        renderLatex(document.getElementById('integral-eq-preview'), buildEquationLatex(window.currentChartData));
    }

    // --- 잔차 그래프 렌더링 ---
    function renderResidualChart(xData, residuals, xLabel) {
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        const textColor = isDark ? '#e2e8f0' : '#1a202c';
        const gridColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)';

        const scatter = {
            x: xData, y: residuals, mode: 'markers', type: 'scatter', name: '잔차',
            marker: { color: 'hsl(195, 75%, 45%)', size: 8, opacity: 0.75 }
        };
        const minX = Math.min(...xData), maxX = Math.max(...xData);
        const zeroLine = {
            x: [minX, maxX], y: [0, 0], mode: 'lines', type: 'scatter', name: '기준선 (잔차=0)',
            line: { color: 'hsl(5, 75%, 55%)', dash: 'dot', width: 2 }
        };

        const layout = {
            autosize: true, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)',
            font: { color: textColor, family: 'Inter' },
            xaxis: { title: xLabel, gridcolor: gridColor, zerolinecolor: gridColor },
            yaxis: { title: '잔차 (실제값 - 예측값)', gridcolor: gridColor, zerolinecolor: gridColor },
            margin: { l: 60, r: 20, t: 15, b: 50 },
            legend: { orientation: "h", y: -0.3 }
        };
        Plotly.newPlot('residual-chart-container', [scatter, zeroLine], layout, { responsive: true, displayModeBar: true, displaylogo: false });
    }

    // --- 예측 계산 (Step 3) ---
    function runPrediction() {
        const inputEl = document.getElementById('predictX');
        if(!inputEl.value) {
            alert("값을 입력해 주세요.");
            return;
        }

        const inputX = parseFloat(inputEl.value);
        if(!window.currentChartData || !window.currentChartData.regResult) {
            alert("먼저 가설 검증 단계(Step 2)에서 분석을 완료해 주세요.");
            return;
        }

        const { regResult, shiftX, xLabel, yLabel } = window.currentChartData;

        let adjX = inputX - shiftX;

        // 예측 수행
        const pred = regResult.predict(adjX);
        const predictedY = pred[1];

        if (isNaN(predictedY) || !isFinite(predictedY)) {
            alert("수학적 한계(예: 로그의 진수가 0 이하)로 인해 해당 값은 예측할 수 없습니다. 모델 범위를 확인해 주세요.");
            return;
        }

        // UI 표시
        const resultBox = document.getElementById('predict-result-box');
        resultBox.classList.remove('hidden');
        document.getElementById('pred-x-name').textContent = xLabel;
        document.getElementById('pred-x-val').textContent = inputX;
        document.getElementById('pred-y-name').textContent = yLabel;
        document.getElementById('pred-y-val').textContent = predictedY.toFixed(2); // 소수점 2자리까지만 표시
    }

    // --- 정적분(Step 4) 관련 계산 ---

    // 다항함수(선형/2~4차) 계수 배열 [최고차항, ..., 상수항] 기준 정적분: F(t2) - F(t1)
    function polynomialDefiniteIntegral(equation, t1, t2) {
        const order = equation.length - 1;
        function antiderivativeAt(t) {
            let sum = 0;
            for (let i = 0; i <= order; i++) {
                const power = order - i;
                const coef = equation[i];
                sum += coef / (power + 1) * Math.pow(t, power + 1);
            }
            return sum;
        }
        return antiderivativeAt(t2) - antiderivativeAt(t1);
    }

    // 심슨(Simpson)의 법칙을 이용한 수치적분 (지수/로그 함수용)
    function simpsonIntegral(f, a, b, n) {
        if (n % 2 === 1) n++;
        const h = (b - a) / n;
        let sum = f(a) + f(b);
        for (let i = 1; i < n; i++) {
            const x = a + i * h;
            sum += (i % 2 === 0 ? 2 : 4) * f(x);
        }
        return (h / 3) * sum;
    }

    function runIntegration() {
        if (!window.currentChartData || !window.currentChartData.regResult) {
            alert("먼저 가설 검증 단계(Step 2)에서 분석을 완료해 주세요.");
            return;
        }
        const inputAEl = document.getElementById('integralA');
        const inputBEl = document.getElementById('integralB');
        if (!inputAEl.value || !inputBEl.value) {
            alert("구간 a, b 값을 모두 입력해주세요.");
            return;
        }
        const a = parseFloat(inputAEl.value);
        const b = parseFloat(inputBEl.value);
        if (a === b) {
            alert("a와 b는 서로 달라야 합니다.");
            return;
        }

        const d = window.currentChartData;
        const { regResult, shiftX, modelType, useZeroAdjust } = d;
        const warnBox = document.getElementById('integral-domain-warning');
        warnBox.classList.add('hidden');

        let sign = 1;
        let lo = a, hi = b;
        if (a > b) { lo = b; hi = a; sign = -1; }

        // 로그 함수는 정의역(x - shift > 0) 확인 필요
        if (modelType === 'logarithmic' && (lo - shiftX) <= 0) {
            warnBox.classList.remove('hidden');
            document.getElementById('integral-domain-warning-text').textContent =
                `로그 함수는 x > ${shiftX} 범위에서만 정의됩니다. 구간 [a, b]를 이 범위 안에서 다시 입력해주세요.`;
            return;
        }

        const f = function(origX) {
            const adjX = origX - shiftX;
            return regResult.predict(adjX)[1];
        };

        const xTerm = getXTermLatex(shiftX, useZeroAdjust);
        let integralValue, methodNote;
        const formulaBox = document.getElementById('integral-formula-box');
        formulaBox.style.display = 'none';

        if (isPolynomialFamily(modelType)) {
            const loT = lo - shiftX, hiT = hi - shiftX;
            integralValue = sign * polynomialDefiniteIntegral(regResult.equation, loT, hiT);
            methodNote = '다항함수이므로 부정적분 공식을 이용해 정확한 값을 계산했습니다.';
            formulaBox.style.display = 'block';
            renderLatex(formulaBox, `${polynomialAntiderivativeLatex(regResult.equation, xTerm)} \\;\\Rightarrow\\; \\int = F(${hi}) - F(${lo})`);
        } else {
            integralValue = sign * simpsonIntegral(f, lo, hi, 200);
            methodNote = `${modelType === 'exponential' ? '지수' : '로그'}함수는 닫힌 형태의 부정적분 대신, 심슨(Simpson)의 법칙을 이용한 수치적분 근사값을 계산했습니다.`;
            formulaBox.style.display = 'block';
            renderLatex(formulaBox, buildEquationLatex(d));
        }

        window.currentIntegralRange = { lo, hi };

        document.getElementById('integral-result-box').classList.remove('hidden');
        document.getElementById('integral-method-note').textContent = methodNote;
        renderLatex(document.getElementById('integral-notation-latex'), `\\int_{${lo}}^{${hi}} f(x)\\,dx ${isPolynomialFamily(modelType) ? '=' : '\\approx'}`);
        document.getElementById('integral-result-val').textContent = integralValue.toFixed(4);

        renderIntegralChart(lo, hi);
    }

    function renderIntegralChart(lo, hi) {
        const d = window.currentChartData;
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        const textColor = isDark ? '#e2e8f0' : '#1a202c';
        const gridColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)';
        const overrides = window.currentChartOptions || {};

        const scatter = {
            x: d.xData, y: d.yData, mode: 'markers', type: 'scatter', name: '실제 데이터',
            marker: { color: overrides.pointColor || 'hsl(255, 65%, 60%)', size: 7, opacity: 0.5 }
        };
        const trendline = {
            x: d.curveX, y: d.curveY, mode: 'lines', type: 'scatter', name: '함수 곡선',
            line: { color: overrides.lineColor || 'hsl(5, 75%, 55%)', dash: 'dash', width: 2 }
        };

        const fillX = [], fillY = [];
        const steps = 100;
        for (let i = 0; i <= steps; i++) {
            const x = lo + (hi - lo) * (i / steps);
            const adjX = x - d.shiftX;
            fillX.push(x);
            fillY.push(d.regResult.predict(adjX)[1]);
        }
        const fillArea = {
            x: fillX, y: fillY, mode: 'lines', type: 'scatter', name: `구간 [${lo}, ${hi}] 넓이`,
            fill: 'tozeroy', fillcolor: 'hsla(255, 65%, 60%, 0.35)', line: { color: 'hsl(255, 65%, 45%)', width: 1 }
        };

        const layout = {
            autosize: true, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)',
            font: { color: textColor, family: 'Inter' },
            title: { text: `구간 [${lo}, ${hi}]에서의 정적분 넓이`, font: { family: 'Outfit', size: 15 } },
            xaxis: { title: d.xLabel, gridcolor: gridColor, zerolinecolor: gridColor },
            yaxis: { title: d.yLabel, gridcolor: gridColor, zerolinecolor: gridColor },
            margin: { l: 55, r: 20, t: 35, b: 50 },
            legend: { orientation: "h", y: -0.2 }
        };

        Plotly.newPlot('integral-chart-container', [scatter, trendline, fillArea], layout, { responsive: true, displayModeBar: true, displaylogo: false });
    }

    // --- 파일 다운로드 ---
    function downloadReport() {
        if(!window.currentChartData) return;
        const d = window.currentChartData;
        const xCol = d.xLabel;
        const yCol = d.yLabel;
        const r2Val = document.getElementById('res-r2').textContent;
        const modelType = document.getElementById('trendModel').options[document.getElementById('trendModel').selectedIndex].text;

        let result_text = `==== 동덕여고 데이터 기반 가설 탐구 결과 레포트 ====\n\n`;
        result_text += `[분석 변수]\n- 독립변수(X): ${xCol}\n- 종속변수(Y): ${yCol}\n\n`;
        result_text += `[분석 모델]\n- ${modelType}\n\n`;
        result_text += `[도출된 수식]\n- y = ${d.equation}\n\n`;
        result_text += `[통계 검증]\n- 설명력 (R²): ${r2Val}\n`;
        result_text += `  (1에 가까울수록 함수가 데이터를 잘 설명함)\n`;

        const blob = new Blob([result_text], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = "가설검증_결과_레포트.txt";
        document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    // --- 샘플 CSV 다운로드 ---
    function downloadSampleCSV() {
        const sampleData = "﻿수면시간,학업성적\n7.5,85\n6,72\n5.5,65\n8,90\n4.5,50\n7,80\n6.5,78\n5,60\n8.5,92\n6,70\n7,82\n5.5,68\n6.5,75\n7.5,88\n5,62\n6,74\n8,86\n4,45\n7,81\n6.5,76\n9,95\n5.8,71\n6.8,80\n7.2,84\n5.2,61\n6.2,73\n8.2,89\n4.8,58\n7.8,87\n6.4,72\n";
        const blob = new Blob([sampleData], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = "sleep_vs_score_example.csv";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
