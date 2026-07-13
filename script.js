// --- 화면 라우팅 (메뉴 전환) ---
    function switchView(viewName) {
        document.getElementById('view-step1').classList.add('hidden');
        document.getElementById('view-step2').classList.add('hidden');
        document.getElementById('view-step3').classList.add('hidden');
        document.getElementById('nav-step1').classList.remove('active');
        document.getElementById('nav-step2').classList.remove('active');
        document.getElementById('nav-step3').classList.remove('active');

        document.getElementById('view-' + viewName).classList.remove('hidden');
        document.getElementById('nav-' + viewName).classList.add('active');
        
        // 예측 뷰로 넘어갈 때 라벨 업데이트
        if(viewName === 'step3' && window.currentChartData) {
            document.getElementById('predict-input-label').textContent = `예측하고 싶은 [${window.currentChartData.xLabel}] 값을 입력하세요:`;
        }
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
        
        if (window.currentChartData) {
            renderChart(
                window.currentChartData.xData, 
                window.currentChartData.yData, 
                window.currentChartData.curveX, 
                window.currentChartData.curveY, 
                window.currentChartData.xLabel, 
                window.currentChartData.yLabel, 
                window.currentChartData.equation
            );
        }
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

    function handleFile(file) {
        if (!file) return;
        const encoding = document.getElementById('encoding-select').value;
        
        // 업로드된 파일명 표시 UI 업데이트
        const icon = document.getElementById('dropzone-icon');
        icon.className = 'fa-solid fa-file-circle-check';
        icon.style.color = 'var(--success)';
        document.getElementById('dropzone-title').innerHTML = `선택된 파일: <span style="color: var(--primary);">${file.name}</span>`;
        document.getElementById('dropzone-desc').textContent = '다른 파일을 분석하려면 다시 클릭하거나 드래그하세요.';

        Papa.parse(file, {
            header: true,
            dynamicTyping: false, // 타입 자동 변환 해제로 파싱 속도 대폭 향상
            worker: true,         // 웹 워커를 사용하여 브라우저 멈춤 방지
            skipEmptyLines: true,
            encoding: encoding,
            complete: function(results) {
                globalData = results.data;
                renderTable(globalData);
                processColumns(globalData);
            },
            error: function(err) {
                alert("파일을 읽는 중 오류가 발생했습니다. 인코딩 설정을 확인해 보세요.");
            }
        });
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

    // --- 통계 분석 (Regression.js + 평행이동 + jStat) ---
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

        // Regression.js 추세선 및 결정계수(R2) 계산 (지수/로그는 일반화 탐색 모델 적용)
        let regResult;
        if(modelType === 'linear') {
            regResult = regression.linear(points, { precision: 4 });
            if(regResult.string.includes('y = ')) regResult.string = regResult.string.replace('y = ', '');
        }
        else if(modelType === 'exponential') {
            regResult = findGeneralizedExponential(points) || regression.exponential(points, { precision: 4 });
            if(regResult.string.includes('y = ')) regResult.string = regResult.string.replace('y = ', '');
        }
        else if(modelType === 'logarithmic') {
            regResult = findGeneralizedLogarithmic(points) || regression.logarithmic(points, { precision: 4 });
            if(regResult.string.includes('y = ')) regResult.string = regResult.string.replace('y = ', '');
        }
        else if(modelType === 'polynomial') {
            regResult = regression.polynomial(points, { order: 2, precision: 4 });
            if(regResult.string.includes('y = ')) regResult.string = regResult.string.replace('y = ', '');
        }

        let equationText = regResult.string;
        
        // 화면에 출력할 수식 보정 (x 대신 평행이동한 (x - shift) 형태 명시)
        if (useZeroAdjust && shiftX !== 0) {
            let shiftStr = shiftX > 0 ? `- ${shiftX}` : `+ ${Math.abs(shiftX)}`;
            equationText = equationText.replace(/x/g, `(x ${shiftStr})`);
        }

        const r2 = regResult.r2;

        // jStat: P-value (항상 선형 피어슨 기준으로 참고용으로만 제공)
        const corr = jStat.corrcoeff(xData, yData);
        const n = xData.length;
        const t_stat = corr * Math.sqrt(n - 2) / Math.sqrt(1 - corr * corr);
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

        // UI 업데이트
        document.getElementById('result-card').classList.remove('hidden');
        document.getElementById('res-eq').textContent = equationText;
        document.getElementById('res-r2').textContent = r2.toFixed(4);
        document.getElementById('res-p').textContent = p_value.toFixed(4);

        const interpBox = document.getElementById('interpretation-box');
        if (r2 >= 0.7) {
            interpBox.innerHTML = `
                <div class="alert alert-success">
                    <i class="fa-solid fa-circle-check"></i>
                    <div>
                        <strong>강력한 설명력!</strong>
                        <p>결정계수(R²)가 ${r2.toFixed(2)}로 매우 높습니다. 선택하신 <b>[${modelType}] 함수식</b>이 두 변수 간의 관계를 훌륭하게 설명하고 있습니다. 탐구 보고서에 이 함수식을 핵심 수학적 근거로 제시해 보세요!</p>
                    </div>
                </div>`;
        } else if (r2 >= 0.4) {
            interpBox.innerHTML = `
                <div class="alert alert-info">
                    <i class="fa-solid fa-chart-bar"></i>
                    <div>
                        <strong>유의미한 설명력</strong>
                        <p>결정계수(R²)가 ${r2.toFixed(2)}로 중간 정도입니다. 선택하신 함수식이 어느 정도 의미 있는 패턴을 보여주지만, 예외적인 데이터(오차)도 제법 존재합니다. 다른 추세선 모델을 적용했을 때 더 높아지는지 비교해 보세요.</p>
                    </div>
                </div>`;
        } else {
            interpBox.innerHTML = `
                <div class="alert alert-warning">
                    <i class="fa-solid fa-triangle-exclamation"></i>
                    <div>
                        <strong>설명력 부족</strong>
                        <p>결정계수(R²)가 ${r2.toFixed(2)}로 많이 낮습니다. 데이터가 선택한 함수식을 거의 따르지 않거나 다른 변수의 영향이 큽니다. 추세선 모델을 다른 함수(다항, 로그 등)로 변경해 보세요!</p>
                    </div>
                </div>`;
        }

        window.currentChartData = { 
            xData, yData, curveX, curveY, 
            xLabel: xCol, yLabel: yCol, 
            equation: equationText,
            regResult: regResult,
            shiftX: shiftX,
            modelType: modelType
        };
        renderChart(xData, yData, curveX, curveY, xCol, yCol, equationText);
    }

    // --- 그래프 렌더링 ---
    function renderChart(xData, yData, curveX, curveY, xLabel, yLabel, equationStr) {
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        const textColor = isDark ? '#e2e8f0' : '#1a202c';
        const gridColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)';
        const paperColor = 'rgba(0,0,0,0)';

        const scatter = {
            x: xData, y: yData, mode: 'markers', type: 'scatter', name: '실제 데이터',
            marker: { color: 'hsl(255, 65%, 60%)', size: 8, opacity: 0.7 }
        };

        const trendline = {
            x: curveX, y: curveY, mode: 'lines', type: 'scatter', name: '예측 곡선',
            line: { color: 'hsl(5, 75%, 55%)', dash: 'dash', width: 3 }
        };

        const layout = {
            autosize: true, plot_bgcolor: paperColor, paper_bgcolor: paperColor,
            font: { color: textColor, family: 'Inter' },
            title: { text: `y = ${equationStr}`, font: { family: 'Outfit', size: 16 } },
            xaxis: { title: xLabel, gridcolor: gridColor, zerolinecolor: gridColor },
            yaxis: { title: yLabel, gridcolor: gridColor, zerolinecolor: gridColor },
            margin: { l: 50, r: 20, t: 40, b: 50 },
            legend: { orientation: "h", y: -0.2 }
        };

        Plotly.newPlot('chart-container', [scatter, trendline], layout, { responsive: true, displayModeBar: true, displaylogo: false });
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

        const { regResult, shiftX, modelType, xLabel, yLabel } = window.currentChartData;
        
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

    // --- 파일 다운로드 ---
    function downloadReport() {
        if(!window.currentChartData) return;
        const xCol = window.currentChartData.xLabel;
        const yCol = window.currentChartData.yLabel;
        const eqVal = document.getElementById('res-eq').textContent;
        const r2Val = document.getElementById('res-r2').textContent;
        const modelType = document.getElementById('trendModel').options[document.getElementById('trendModel').selectedIndex].text;

        let result_text = `==== 동덕여고 데이터 기반 가설 탐구 결과 레포트 ====\n\n`;
        result_text += `[분석 변수]\n- 독립변수(X): ${xCol}\n- 종속변수(Y): ${yCol}\n\n`;
        result_text += `[분석 모델]\n- ${modelType}\n\n`;
        result_text += `[도출된 수식]\n- y = ${eqVal}\n\n`;
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
        const sampleData = "\uFEFF수면시간,학업성적\n7.5,85\n6,72\n5.5,65\n8,90\n4.5,50\n7,80\n6.5,78\n5,60\n8.5,92\n6,70\n7,82\n5.5,68\n6.5,75\n7.5,88\n5,62\n6,74\n8,86\n4,45\n7,81\n6.5,76\n9,95\n5.8,71\n6.8,80\n7.2,84\n5.2,61\n6.2,73\n8.2,89\n4.8,58\n7.8,87\n6.4,72\n";
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