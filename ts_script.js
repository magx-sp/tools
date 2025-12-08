document.addEventListener('DOMContentLoaded', () => {
        const widthInput = document.getElementById('width-input');
        const heightInput = document.getElementById('height-input');
        const radiusInput = document.getElementById('radius-input');
        const previewArea = document.getElementById('preview-area');
        const pdfDownloadBtn = document.getElementById('pdf-download-btn');

        // --- Constants ---
        const BLEED = 3; // 塗り足し幅 (3mm)
        const SAFE_MARGIN = 2;
        const PT_TO_MM = 0.352778;
        
        const TOMBO_LINE_WIDTH_PT = 0.3;
        const OTHER_LINE_WIDTH_PT = 0.5;
        const TOMBO_LINE_WIDTH_MM = TOMBO_LINE_WIDTH_PT * PT_TO_MM;
        const OTHER_LINE_WIDTH_MM = OTHER_LINE_WIDTH_PT * PT_TO_MM;

        // 破線パターン (2pt = 0.7056mm)
        const DASH_LENGTH_MM = 2 * PT_TO_MM; 

        // ** コーナートンボの固定サイズ **
        const SHORT_LINE = 9; // 短い線 (9mm)
        const LONG_LINE = 12; // 長い線 (12mm)
        const MARK_AREA_SIZE = 12; // アートボードマージン
        
        // ** センタートンボの固定サイズ **
        const CENTER_H_LONG = 25.4; 
        const CENTER_V_SHORT = 25.4; 
        const CENTER_H_SHORT = 8.467;
        const CENTER_V_LONG = 8.467; 
        
        // ** センタートンボの最終位置 (塗り足し線からのオフセット) **
        const LONG_LINE_OFFSET = 3.35; // 25.4mm線のオフセット
        const SHORT_LINE_OFFSET = 1.23; // 8.467mm線のオフセット

        let currentData = {};

        function updatePreview() {
            let width = parseFloat(widthInput.value) || 0;
            let height = parseFloat(heightInput.value) || 0;
            let radius = parseFloat(radiusInput.value) || 0;
            
            if (width < 1 || height < 1) {
                previewArea.innerHTML = '<p style="color: #888;">有効なサイズを入力してください。</p>';
                return;
            }
            // ❗ 角丸の入力値上書き処理を削除 (表示値はそのまま維持)
            radius = Math.max(0, Math.min(radius, width / 2, height / 2));
            
            currentData = { width, height, radius };
            
            const artboardWidth = width + MARK_AREA_SIZE * 2;
            const artboardHeight = height + MARK_AREA_SIZE * 2;
            
            const svg = generateSVG(width, height, radius, artboardWidth, artboardHeight);
            previewArea.innerHTML = '';
            previewArea.appendChild(svg);
        }

        function generateSVG(width, height, radius, artboardWidth, artboardHeight) {
            const svgNS = "http://www.w3.org/2000/svg";
            const svg = document.createElementNS(svgNS, "svg");
            
            svg.setAttribute("viewBox", `0 0 ${artboardWidth} ${artboardHeight}`);
            
            const style = document.createElementNS(svgNS, "style");
            style.textContent = `
                .guide-line { stroke-width: ${OTHER_LINE_WIDTH_MM}; }
                .tombo-line { stroke: var(--preview-tombo-color); stroke-width: ${TOMBO_LINE_WIDTH_MM}; }
            `;
            svg.appendChild(style);

            const g = document.createElementNS(svgNS, "g");
            g.setAttribute("fill", "none");

            const baseX = MARK_AREA_SIZE;
            const baseY = MARK_AREA_SIZE;
            const centerX = baseX + width / 2;
            const centerY = baseY + height / 2;

            // --- 1. ガイド線 (プレビュー) ---
            g.innerHTML = `
                <rect class="guide-line" x="${baseX - BLEED}" y="${baseY - BLEED}" width="${width + BLEED * 2}" height="${height + BLEED * 2}" rx="${radius + BLEED}" stroke="var(--preview-bleed-color)" stroke-dasharray="${DASH_LENGTH_MM} ${DASH_LENGTH_MM}"/>
                <rect class="guide-line" x="${baseX}" y="${baseY}" width="${width}" height="${height}" rx="${radius}" stroke="var(--preview-finish-color)"/>
                ${(width > SAFE_MARGIN * 2 && height > SAFE_MARGIN * 2) ? 
                    `<rect class="guide-line" x="${baseX + SAFE_MARGIN}" y="${baseY + SAFE_MARGIN}" width="${width - SAFE_MARGIN * 2}" height="${height - SAFE_MARGIN * 2}" rx="${Math.max(0, radius - SAFE_MARGIN)}" stroke="var(--preview-safe-color)" stroke-dasharray="${DASH_LENGTH_MM} ${DASH_LENGTH_MM}"/>`
                    : ''}
            `;

            const tomboGroup = document.createElementNS(svgNS, "g");
            tomboGroup.setAttribute("class", "tombo-line");
            
            // --- 2. コーナートンボ (連結パス) ---
            const corners = [
                {x: baseX, y: baseY}, {x: baseX + width, y: baseY},
                {x: baseX, y: baseY + height}, {x: baseX + width, y: baseY + height}
            ];
            
            corners.forEach((corner, i) => {
                const signX = (i % 2 === 0) ? -1 : 1;
                const signY = (i < 2) ? -1 : 1;
                
                // L字パス 1: 長い水平線 (12mm) + 短い垂直線 (9mm)
                const L1_start_X = corner.x + signX * LONG_LINE;
                const L1_start_Y = corner.y + signY * BLEED;
                const L1_end_Y = corner.y + signY * (BLEED + SHORT_LINE);

                tomboGroup.innerHTML += `<path d="M ${L1_start_X},${L1_start_Y} L ${corner.x},${L1_start_Y} L ${corner.x},${L1_end_Y}" />`;

                // L字パス 2: 短い水平線 (9mm) + 長い垂直線 (12mm)
                const L2_start_X = corner.x + signX * (BLEED + SHORT_LINE);
                const L2_start_Y = corner.y;
                const L2_corner_X = corner.x + signX * BLEED;
                const L2_end_Y = corner.y + signY * LONG_LINE;

                tomboGroup.innerHTML += `<path d="M ${L2_start_X},${L2_start_Y} L ${L2_corner_X},${L2_start_Y} L ${L2_corner_X},${L2_end_Y}" />`;
            });

            // --- 3. センタートンボ (修正位置) ---
            
            // 上下 - 水平線 (長い線 25.4mm)
            const topY_H = baseY - BLEED - LONG_LINE_OFFSET;
            const bottomY_H = baseY + height + BLEED + LONG_LINE_OFFSET;
            tomboGroup.innerHTML += `<path d="M ${centerX - CENTER_H_LONG / 2},${topY_H} h ${CENTER_H_LONG}" />`;
            tomboGroup.innerHTML += `<path d="M ${centerX - CENTER_H_LONG / 2},${bottomY_H} h ${CENTER_H_LONG}" />`;
            
            // 上下 - 垂直線 (短い線 8.467mm)
            const topY_V_start = baseY - BLEED - SHORT_LINE_OFFSET;
            const bottomY_V_start = baseY + height + BLEED + SHORT_LINE_OFFSET;
            tomboGroup.innerHTML += `<path d="M ${centerX},${topY_V_start} v -${CENTER_V_LONG}" />`; 
            tomboGroup.innerHTML += `<path d="M ${centerX},${bottomY_V_start} v ${CENTER_V_LONG}" />`; 

            // 左右 - 垂直線 (長い線 25.4mm)
            const leftX_V = baseX - BLEED - LONG_LINE_OFFSET;
            const rightX_V = baseX + width + BLEED + LONG_LINE_OFFSET;
            tomboGroup.innerHTML += `<path d="M ${leftX_V},${centerY - CENTER_V_SHORT / 2} v ${CENTER_V_SHORT}" />`;
            tomboGroup.innerHTML += `<path d="M ${rightX_V},${centerY - CENTER_V_SHORT / 2} v ${CENTER_V_SHORT}" />`;
            
            // 左右 - 水平線 (短い線 8.467mm)
            const leftX_H_start = baseX - BLEED - SHORT_LINE_OFFSET;
            const rightX_H_start = baseX + width + BLEED + SHORT_LINE_OFFSET;
            tomboGroup.innerHTML += `<path d="M ${leftX_H_start},${centerY} h -${CENTER_H_SHORT}" />`; 
            tomboGroup.innerHTML += `<path d="M ${rightX_H_start},${centerY} h ${CENTER_H_SHORT}" />`; 

            g.appendChild(tomboGroup);
            svg.appendChild(g);
            return svg;
        }

        function handlePDFDownload() {
            const { jsPDF } = window.jspdf;
            const { width, height, radius } = currentData;

            const artboardWidth = width + MARK_AREA_SIZE * 2;
            const artboardHeight = height + MARK_AREA_SIZE * 2;
            
            const doc = new jsPDF({
                orientation: artboardWidth > artboardHeight ? 'l' : 'p',
                unit: 'mm',
                format: [artboardWidth, artboardHeight]
            });

            const baseX = MARK_AREA_SIZE;
            const baseY = MARK_AREA_SIZE;
            const centerX = baseX + width / 2;
            const centerY = baseY + height / 2;

            // --- 1. ガイド線（CMYK指定） ---
            doc.setLineWidth(OTHER_LINE_WIDTH_MM);
            
            // 塗り足し線 (Bleed): C100 M0 Y100 K0
            doc.setDrawColor(100, 0, 100, 0); 
            doc.setLineDashPattern([DASH_LENGTH_MM, DASH_LENGTH_MM], 0);
            doc.roundedRect(baseX - BLEED, baseY - BLEED, width + BLEED * 2, height + BLEED * 2, radius + BLEED, radius + BLEED, 'S');
            
            // 仕上がりサイズ (Finish): C0 M100 Y0 K0 (M100)
            doc.setDrawColor(0, 100, 0, 0); 
            doc.setLineDashPattern([], 0);
            doc.roundedRect(baseX, baseY, width, height, radius, radius, 'S');
            
            // 安全範囲 (Safe Margin): C100 M0 Y0 K0
            if (width > SAFE_MARGIN * 2 && height > SAFE_MARGIN * 2) {
                doc.setDrawColor(100, 0, 0, 0);
                doc.setLineDashPattern([DASH_LENGTH_MM, DASH_LENGTH_MM], 0);
                let safeRadius = Math.max(0, radius - SAFE_MARGIN);
                doc.roundedRect(baseX + SAFE_MARGIN, baseY + SAFE_MARGIN, width - SAFE_MARGIN * 2, height - SAFE_MARGIN * 2, safeRadius, safeRadius, 'S');
            }

            // --- 2. トンボ: レジストレーションカラー (C100 M100 Y100 K100) ---
            doc.setDrawColor(100, 100, 100, 100); 
            doc.setLineWidth(TOMBO_LINE_WIDTH_MM);
            doc.setLineDashPattern([], 0);
            
            // ** 2-1. コーナートンボ (連結パス) **
            const corners = [
                {x: baseX, y: baseY}, {x: baseX + width, y: baseY},
                {x: baseX, y: baseY + height}, {x: baseX + width, y: baseY + height}
            ];

            corners.forEach((corner, i) => {
                const signX = (i % 2 === 0) ? -1 : 1;
                const signY = (i < 2) ? -1 : 1;
                
                // L字パス 1: 長い水平線 (12mm) + 短い垂直線 (9mm)
                const L1_start_X = corner.x + signX * LONG_LINE;
                const L1_start_Y = corner.y + signY * BLEED;
                const L1_moves = [
                    [-signX * LONG_LINE, 0], 
                    [0, signY * SHORT_LINE]  
                ];
                doc.lines(L1_moves, L1_start_X, L1_start_Y); 

                // L字パス 2: 短い水平線 (9mm) + 長い垂直線 (12mm)
                const L2_start_X = corner.x + signX * (BLEED + SHORT_LINE);
                const L2_start_Y = corner.y;
                const L2_moves = [
                    [-signX * SHORT_LINE, 0], 
                    [0, signY * LONG_LINE]    
                ];
                doc.lines(L2_moves, L2_start_X, L2_start_Y);
            });

            // ** 2-2. センタートンボ (修正位置) **

            // 上下 - 水平線 (長い線 25.4mm)
            const topY_H = baseY - BLEED - LONG_LINE_OFFSET;
            const bottomY_H = baseY + height + BLEED + LONG_LINE_OFFSET;
            doc.lines([[CENTER_H_LONG, 0]], centerX - CENTER_H_LONG / 2, topY_H);
            doc.lines([[CENTER_H_LONG, 0]], centerX - CENTER_H_LONG / 2, bottomY_H);
            
            // 上下 - 垂直線 (短い線 8.467mm)
            const topY_V_start = baseY - BLEED - SHORT_LINE_OFFSET;
            const bottomY_V_start = baseY + height + BLEED + SHORT_LINE_OFFSET;
            doc.lines([[0, -CENTER_V_LONG]], centerX, topY_V_start);
            doc.lines([[0, CENTER_V_LONG]], centerX, bottomY_V_start);

            // 左右 - 垂直線 (長い線 25.4mm)
            const leftX_V = baseX - BLEED - LONG_LINE_OFFSET;
            const rightX_V = baseX + width + BLEED + LONG_LINE_OFFSET;
            doc.lines([[0, CENTER_V_SHORT]], leftX_V, centerY - CENTER_V_SHORT / 2);
            doc.lines([[0, CENTER_V_SHORT]], rightX_V, centerY - CENTER_V_SHORT / 2);
            
            // 左右 - 水平線 (短い線 8.467mm)
            const leftX_H_start = baseX - BLEED - SHORT_LINE_OFFSET;
            const rightX_H_start = baseX + width + BLEED + SHORT_LINE_OFFSET;
            doc.lines([[-CENTER_H_SHORT, 0]], leftX_H_start, centerY);
            doc.lines([[CENTER_H_SHORT, 0]], rightX_H_start, centerY);
            
            // ファイル名もバージョン番号を上げて保存
            doc.save(`template_${currentData.width}x${currentData.height}_final-v28.pdf`);
        }

        [widthInput, heightInput, radiusInput].forEach(input => input.addEventListener('input', updatePreview));
        pdfDownloadBtn.addEventListener('click', handlePDFDownload);

        updatePreview();
    });