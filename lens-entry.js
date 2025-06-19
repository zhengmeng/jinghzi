document.addEventListener('DOMContentLoaded', () => {
    const lensEntryForm = document.getElementById('lens-entry-form');
    const addLensButton = document.getElementById('add-lens');
    const batchUploadButton = document.getElementById('batch-upload');
    const modal = document.getElementById('modal');
    const batchSubmitButton = document.getElementById('batch-submit');
    const batchProductNameSelect = document.getElementById('batch-product-name');
    const resultsDiv = document.getElementById('results');
    const viewQrcodesButton = document.getElementById('view-qrcodes');
    const printAllQrcodesButton = document.getElementById('print-all-qrcodes');
    const excelButton = document.getElementById('excel');
    const grid = document.getElementById('batch-grid');
    let lensCounter = 1;
    let isDragging = false;
    let startCell = null;
    let lastCell = null; // 用于跟踪最后一个经过的单元格
    let isRightClick = false;

    const lensData = {
        '爱眼星点矩阵1.60MR-8近视管理镜片': { refraction: '1.598', abbeNumber: '41' },
        '爱眼星点矩阵1.67近视管理镜片': { refraction: '1.667', abbeNumber: '32' },
        '爱眼星1.56多微透镜离焦专业版': { refraction: '1.564', abbeNumber: '38' },
        '爱眼星1.60双面复合多微透镜离焦镜片': { refraction: '1.595', abbeNumber: '32' },
    };

    const productNameCode = {
        '爱眼星点矩阵1.60MR-8近视管理镜片': 'A',
        '爱眼星点矩阵1.67近视管理镜片': 'B',
        '爱眼星1.56多微透镜离焦专业版': 'C',
        '爱眼星1.60双面复合多微透镜离焦镜片': 'D',
    };

    const technicalParams = {
        '爱眼星点矩阵1.60MR-8近视管理镜片': {
            dia: [
                { powerRange: { min: 0.25, max: 6.00 }, cylinderRange: { min: -4.00, max: 0.00 }, dia: '65mm' },
                { powerRange: { min: -6.00, max: 0.00 }, cylinderRange: { min: -4.00, max: 0.00 }, dia: '75mm' },
                { powerRange: { min: -8.00, max: -6.25 }, cylinderRange: { min: -4.00, max: 0.00 }, dia: '70mm' },
                { powerRange: { min: -10.00, max: -8.25 }, cylinderRange: { min: -2.00, max: 0.00 }, dia: '70mm' },
            ],
            thickness: [
                { powerRange: { min: 0.25, max: 0.75 }, cylinderRange: { min: -4.00, max: 0.00 }, et: '1.5mm' },
                { powerRange: { min: 1.00, max: 1.50 }, cylinderRange: { min: -4.00, max: 0.00 }, et: '1.3mm' },
                { powerRange: { min: 1.75, max: 6.00 }, cylinderRange: { min: -4.00, max: 0.00 }, et: '1.1mm' },
                { powerRange: { min: -0.25, max: 0.00 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '2.0mm' },
                { powerRange: { min: -0.50, max: -0.50 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.8mm' },
                { powerRange: { min: -0.75, max: -0.75 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.6mm' },
                { powerRange: { min: -1.00, max: -1.00 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.4mm' },
                { powerRange: { min: -10.00, max: -1.25 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.2mm' },
            ],
        },
        '爱眼星点矩阵1.67近视管理镜片': {
            dia: [
                { powerRange: { min: 0.25, max: 6.00 }, cylinderRange: { min: -2.00, max: 0.00 }, dia: '65mm' },
                { powerRange: { min: -6.00, max: 0.00 }, cylinderRange: { min: -4.00, max: 0.00 }, dia: '75mm' },
                { powerRange: { min: -12.00, max: -6.25 }, cylinderRange: { min: -4.00, max: 0.00 }, dia: '70mm' },
                { powerRange: { min: -8.00, max: 0.00 }, cylinderRange: { min: -6.00, max: -4.25 }, dia: '70mm' },
            ],
            thickness: [
                { powerRange: { min: 0.25, max: 0.75 }, cylinderRange: { min: -4.00, max: 0.00 }, et: '1.5mm' },
                { powerRange: { min: 1.00, max: 1.50 }, cylinderRange: { min: -4.00, max: 0.00 }, et: '1.3mm' },
                { powerRange: { min: 1.75, max: 6.00 }, cylinderRange: { min: -4.00, max: 0.00 }, et: '1.1mm' },
                { powerRange: { min: -0.50, max: -0.25 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.9mm' },
                { powerRange: { min: -0.75, max: -0.75 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.6mm' },
                { powerRange: { min: -1.00, max: -1.00 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.4mm' },
                { powerRange: { min: -1.25, max: -1.25 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.2mm' },
                { powerRange: { min: -15.00, max: -1.50 }, cylinderRange: { min: -4.00, max: 0.00 }, ct: '1.15mm' },
            ],
        },
        '爱眼星1.56多微透镜离焦专业版': {
            dia: [
                { powerRange: { min: 0.00, max: 0.00 }, dia: '75mm', ct: '2.0±0.3' },
                { powerRange: { min: -1.00, max: -0.25 }, dia: '75mm', ct: '1.8±0.3' },
                { powerRange: { min: -2.00, max: -1.25 }, dia: '75mm', ct: '1.6±0.3' },
                { powerRange: { min: -6.00, max: -2.25 }, dia: '75mm', ct: '1.4±0.3' },
                { powerRange: { min: -10.00, max: -6.25 }, dia: '70mm', ct: '1.2±0.3' },
            ],
            thickness: [
                { powerRange: { min: 0.00, max: 0.00 }, ct: '2.0±0.3' },
                { powerRange: { min: -1.00, max: -0.25 }, ct: '1.8±0.3' },
                { powerRange: { min: -2.00, max: -1.25 }, ct: '1.6±0.3' },
                { powerRange: { min: -6.00, max: -2.25 }, ct: '1.4±0.3' },
                { powerRange: { min: -10.00, max: -6.25 }, ct: '1.2±0.3' },
            ]
        },
        '爱眼星1.60双面复合多微透镜离焦镜片': {
            dia: [
                { powerRange: { min: 0.00, max: 0.00 }, dia: '75' },
                { powerRange: { min: -1.00, max: -0.25 }, dia: '75' },
                { powerRange: { min: -2.00, max: -1.25 }, dia: '75' },
                { powerRange: { min: -6.00, max: -2.25 }, dia: '75' },
                { powerRange: { min: -12.00, max: -6.25 }, dia: '72' },
            ],
            thickness: [
                { powerRange: { min: 0.00, max: 0.00 }, ct: '2.1±0.3' },
                { powerRange: { min: -1.00, max: -0.25 }, ct: '1.8±0.3' },
                { powerRange: { min: -2.00, max: -1.25 }, ct: '1.6±0.3' },
                { powerRange: { min: -6.00, max: -2.25 }, ct: '1.4±0.3' },
                { powerRange: { min: -12.00, max: -6.25 }, ct: '1.2±0.3' },
            ]
        }
    };

    const generateUniqueSerialNumber = (baseSerialNumber, index) => {
        return `${baseSerialNumber}-${index + 1}`;
    };

    const generateLensData = (productName, spherical, cylinder, quantity) => {
        const selectedData = lensData[productName];
        const lenses = [];
        const baseSerialNumber = new Date().toISOString().replace(/[-:.TZ]/g, '');

        for (let i = 0; i < quantity; i++) {
            const serialNumber = generateUniqueSerialNumber(baseSerialNumber, i);
            lenses.push({
                serial_number: serialNumber,
                product_name: productName,
                spherical: parseFloat(spherical).toFixed(2),
                cylinder: parseFloat(cylinder).toFixed(2),
                refraction: selectedData ? selectedData.refraction : '',
                abbe_number: selectedData ? selectedData.abbeNumber : '',
                dia: '75mm', // 默认值，稍后根据参数更新
                ct: '2.0±0.3mm', // 默认值，稍后根据参数更新
                production_date: new Date().toISOString().split('T')[0], // 生产日期默认当天
                grade: 'A级',
            });
        }

        return lenses;
    };

    const formatLensValue = (value) => {
        return parseFloat(value).toFixed(2);
    };

    const updateLensParams = (lens) => {
        const spherical = parseFloat(lens.spherical);
        const cylinder = parseFloat(lens.cylinder);
        const paramsDia = technicalParams[lens.product_name].dia.find(param => {
            return spherical >= param.powerRange.min && spherical <= param.powerRange.max &&
                   (!param.cylinderRange || (cylinder >= param.cylinderRange.min && cylinder <= param.cylinderRange.max));
        });

        const paramsThickness = technicalParams[lens.product_name].thickness.find(param => {
            return spherical >= param.powerRange.min && spherical <= param.powerRange.max &&
                   (!param.cylinderRange || (cylinder >= param.cylinderRange.min && cylinder <= param.cylinderRange.max));
        });

        if (paramsDia) {
            lens.dia = paramsDia.dia;
        }

        if (paramsThickness) {
            if (paramsThickness.ct) {
                lens.ct = paramsThickness.ct;
                lens.et = undefined; // 确保 et 不会干扰 ct
            }
            if (paramsThickness.et) {
                lens.et = paramsThickness.et;
                lens.ct = undefined; // 确保 ct 不会干扰 et
            }
        } else {
            console.error(`No technical parameters found for spherical: ${spherical}, cylinder: ${cylinder} in product: ${lens.product_name}`);
        }

        console.log('Updated lens params:', lens); // 调试信息
        return lens;
    };

    const generatePrintContent = (lens, callback) => {
        lens = updateLensParams(lens);

        const { serial_number, product_name, spherical, cylinder, refraction, abbe_number, dia, ct, et, production_date, grade } = lens;
        const productCode = productNameCode[product_name] || '';
        const sphericalFormatted = (spherical > 0 ? `+${formatLensValue(spherical)}` : formatLensValue(spherical));
        const cylinderFormatted = (cylinder > 0 ? `+${formatLensValue(cylinder)}` : formatLensValue(cylinder));
        const positiveSpherical = (parseFloat(spherical) + parseFloat(cylinder) > 0 ? `+${formatLensValue(parseFloat(spherical) + parseFloat(cylinder))}` : formatLensValue(parseFloat(spherical) + parseFloat(cylinder)));
        const positiveCylinder = (parseFloat(-cylinder) > 0 ? `+${formatLensValue(-cylinder)}` : formatLensValue(-cylinder));

        // 生产日期格式转换
        const formattedProductionDate = production_date.replace(/-/g, '');

        const qrContent = `http://plmsys.aiforoptometry.com/detail?serial_number=${encodeURIComponent(serial_number)}`;

        const scale = 2; // 缩放比例，设置为2倍分辨率
        const canvasWidth = 375 * scale; // 实际宽度乘以缩放比例
        const canvasHeight = 250 * scale; // 实际高度乘以缩放比例
        const canvas = document.createElement('canvas');
        canvas.width = canvasWidth;
        canvas.height = canvasHeight;
        const ctx = canvas.getContext('2d');

        const headerHeight = canvasHeight / 3;
        const contentHeight = (canvasHeight / 3) * 2;
        const qrSize = contentHeight * 0.6; // 二维码大小为参数区域高度的60%
        const qrX = canvasWidth - qrSize - 10 * scale; // 二维码在画布右下角
        const qrY = canvasHeight - qrSize - 10 * scale; // 二维码在画布右下角

        const qrCanvas = document.createElement('canvas');
        qrCanvas.width = qrSize;
        qrCanvas.height = qrSize;

        QRCode.toCanvas(qrCanvas, qrContent, { width: qrSize, height: qrSize, errorCorrectionLevel: 'H' }, (err) => {
            if (err) {
                console.error('QR Code generation error:', err);
                return;
            }

            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(qrCanvas, qrX, qrY, qrSize, qrSize);

            const fontSize = canvasWidth * 0.04; // 增加字体大小为7%的宽度
            const productNameFontSize = fontSize * 1.20; // 品名字体稍大一些
            const lineHeight = fontSize * 1.5; // 行高

            // 绘制虚线
            ctx.setLineDash([10, 5]);
            ctx.lineWidth = 4;
            ctx.beginPath();
            ctx.moveTo(0, headerHeight);
            ctx.lineTo(canvasWidth, headerHeight);
            ctx.stroke();
            ctx.setLineDash([]); // 清除虚线设置

            // 绘制文本框，上部贴近画布边缘
            const textBoxMargin = 10 * scale; // 文本框边距
            const textBoxX = textBoxMargin;
            const textBoxY = textBoxMargin - 20;
            const textBoxWidth = canvasWidth - textBoxMargin * 2;
            const textBoxHeight = headerHeight - textBoxMargin * 2;
            ctx.strokeRect(textBoxX, textBoxY, textBoxWidth, textBoxHeight);

            // 旋转文本框中的文字180°
            ctx.save();
            ctx.translate(textBoxX + textBoxWidth / 2, textBoxY + textBoxHeight / 2);
            ctx.rotate(Math.PI);
            ctx.translate(-textBoxX - textBoxWidth / 2, -textBoxY - textBoxHeight / 2);

            ctx.font = `900 ${fontSize}px SimHei`; // 使用黑体并加粗
            ctx.textAlign = 'left';

            // 文本框内的文本
            const textXOffset = textBoxX + 10; // 调整X轴偏移量
            let textYOffset = textBoxY - 15 + lineHeight; // 调整Y轴偏移量

            // 品名和虚线之间一个行距
            ctx.fillText(`品名: ${product_name}`, textXOffset, textYOffset);

            // 球镜S和柱镜C
            textYOffset += lineHeight;
            ctx.font = `900 ${fontSize * 1.3}px SimHei`; // 放大字体 1.2 倍
            ctx.fillText(`球镜S: ${sphericalFormatted}   柱镜C: ${cylinderFormatted}`, textXOffset, textYOffset);

            // 正柱镜格式的度数，前缩进2个汉字的距离
            textYOffset += lineHeight;
            ctx.fillText(`    S: ${positiveSpherical}       C: ${positiveCylinder}`, textXOffset, textYOffset);

            // 恢复字体大小以继续绘制其他文本
            ctx.font = `900 ${fontSize}px SimHei`;
            ctx.restore();


            ctx.font = `900 ${productNameFontSize}px SimHei`;
            ctx.fillText(`品名: ${product_name}`, canvasWidth * 0.05, headerHeight + lineHeight * 2);

            // 下部分内容整体上移半行距离
            const tableX = canvasWidth * 0.05;
            const tableY = headerHeight + lineHeight * 3.7;
            const cellWidth = 170 * scale;
            const cellHeight = lineHeight * 1.5;
            const tableData = [
                [`折射率:${refraction}`, `阿贝数:${abbe_number}`],
                [et ? `ET: ${et}` : `CT:${ct}`, `Dia:${dia}`],
                [`生产日期:${formattedProductionDate}`, `等级:${grade}`]
            ];

            tableData.forEach((row, rowIndex) => {
                row.forEach((cell, cellIndex) => {
                    ctx.fillText(cell, tableX + cellIndex * cellWidth, tableY + rowIndex * cellHeight);
                });
            });

            const img = new Image();
            img.src = canvas.toDataURL();
            img.width = 375; // 设置显示宽度
            img.height = 250; // 设置显示高度
            img.style.margin = '5px';

            console.log('Image generated for modal:', img);

            callback(img);
        });
    };

    // 现有表单提交功能
    if (lensEntryForm) {
        lensEntryForm.addEventListener('submit', (e) => {
            e.preventDefault();

            const entries = lensEntryForm.querySelectorAll('.lens-entry');
            let lenses = [];

            entries.forEach((entry, index) => {
                const productName = entry.querySelector(`[name="product-name"]`).value;
                const spherical = entry.querySelector(`[name="spherical"]`).value;
                const cylinder = entry.querySelector(`[name="cylinder"]`).value;
                const quantity = parseInt(entry.querySelector(`[name="quantity"]`).value, 10);

                const lensBatch = generateLensData(productName, spherical, cylinder, quantity);
                lenses = lenses.concat(lensBatch);

                lensBatch.forEach(lens => {
                    generatePrintContent(lens, (img) => {
                        resultsDiv.appendChild(img);
                    });
                });
            });

            console.log('录入的镜片信息:', lenses);

            // 移除后端提交逻辑，直接提示提交成功
            alert(`提交成功，共录入 ${lenses.length} 片镜片`);
            console.log('镜片数据已在本地保存:', lenses);

            entries.forEach((entry) => {
                lensEntryForm.removeChild(entry);
            });

            lensEntryForm.reset();
            lensCounter = 1;
            viewQrcodesButton.style.display = 'inline-block';
        });
    }

    // 动态生成产品名选项
    const productNameSelect = document.getElementById('product-name-0');
    for (const productName in lensData) {
        const option = document.createElement('option');
        option.value = productName;
        option.textContent = productName;
        productNameSelect.appendChild(option);
    }

    for (const productName in lensData) {
        const option = document.createElement('option');
        option.value = productName;
        option.textContent = productName;
        batchProductNameSelect.appendChild(option);
    }

    // 新增条目功能
    if (addLensButton) {
        addLensButton.addEventListener('click', () => {
            const newLensEntry = document.createElement('div');
            newLensEntry.classList.add('lens-entry');
            newLensEntry.innerHTML = `
                <label for="product-name-${lensCounter}">品名:</label>
                <select id="product-name-${lensCounter}" name="product-name" required>
                    ${Object.keys(lensData).map(productName => `<option value="${productName}">${productName}</option>`).join('')}
                </select>
                <label for="spherical-${lensCounter}">球镜度数:</label>
                <input type="text" id="spherical-${lensCounter}" name="spherical" required>
                <label for="cylinder-${lensCounter}">柱镜度数:</label>
                <input type="text" id="cylinder-${lensCounter}" name="cylinder" required>
                <label for="quantity-${lensCounter}">数量:</label>
                <input type="number" id="quantity-${lensCounter}" name="quantity" value="1" min="1" required>
                <button type="button" class="remove-lens">删除</button>
            `;
            lensEntryForm.insertBefore(newLensEntry, addLensButton);
            lensCounter++;
        });
    }

    // 删除条目功能
    lensEntryForm.addEventListener('click', (e) => {
        if (e.target && e.target.classList.contains('remove-lens')) {
            const entryToRemove = e.target.closest('.lens-entry');
            if (entryToRemove) {
                lensEntryForm.removeChild(entryToRemove);
            }
        }
    });

    // 批量录入模态窗口
    if (batchUploadButton) {
        batchUploadButton.addEventListener('click', () => {
            const batchModal = new bootstrap.Modal(modal);
            batchModal.show();
        });
    }

    // 批量录入提交功能
    if (batchSubmitButton) {
        batchSubmitButton.addEventListener('click', () => {
            const selectedProductName = batchProductNameSelect.value;
            const cells = document.querySelectorAll('.batch-cell');
            let lenses = [];

            cells.forEach(cell => {
                const quantity = parseInt(cell.value || '0', 10);
                if (quantity > 0) {
                    const { spherical, cylinder } = cell.dataset;
                    const lensBatch = generateLensData(selectedProductName, spherical, cylinder, quantity);
                    lenses = lenses.concat(lensBatch);

                    lensBatch.forEach(lens => {
                        generatePrintContent(lens, (img) => {
                            resultsDiv.appendChild(img);
                        });
                    });
                }
            });

            console.log('批量录入的镜片信息:', lenses);
            const batchModal = bootstrap.Modal.getInstance(modal);
            batchModal.hide();

            // 清空所有单元格内容（不清空标签行和标签列）
            cells.forEach(cell => {
                cell.value = '';
                cell.classList.remove('selected');
            });

            batchProductNameSelect.value = '';

            // 修改批量录入提交逻辑
            alert(`提交成功，共录入 ${lenses.length} 片镜片`);
            console.log('批量录入的镜片数据已在本地保存:', lenses);
        });
    }

    if (excelButton) {
        excelButton.addEventListener('click', () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.xlsx, .xls';
            input.addEventListener('change', (e) => {
                const file = e.target.files[0];
                const reader = new FileReader();
                reader.onload = function (event) {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    // 去掉首行和首列
                    const filteredData = jsonData.slice(1).map(row => row.slice(1));

                    // 打印数据到控制台
                    console.log(filteredData);

                    // 动态填充到棋盘格
                    filteredData.forEach((row, rowIndex) => {
                        row.forEach((value, colIndex) => {
                            const cell = document.querySelector(
                                `.batch-cell[data-spherical="${(-rowIndex * 0.25).toFixed(2)}"]` +
                                `[data-cylinder="${(-colIndex * 0.25).toFixed(2)}"]`
                            );
                            if (cell) {
                                cell.value = value !== null ? value : ''; // 将值填入单元格
                            }
                        });
                    });
                };
                reader.readAsArrayBuffer(file);
            });
            input.click();
        });
    }

    // 创建批量录入的棋盘格
    const createBatchGrid = () => {
        const rows = 45; // 总行数
        const cols = 21; // 标签列之外的总列数

        // 清空现有内容
        grid.innerHTML = '';

        // 创建第一行（列标签，Cylinder方向）
        for (let col = 0; col <= cols; col++) {
            const colLabel = document.createElement('div');
            colLabel.classList.add('axis-label');
            if (col > 0) {
                colLabel.textContent = (0 - (col - 1) * 0.25).toFixed(2);
            }
            grid.appendChild(colLabel);
        }

        // 创建其余行（每行的第一个单元格为行标签，Spherical方向）
        for (let row = 0; row < rows; row++) {
            // 创建每行的首个标签（行标签）
            const rowLabel = document.createElement('div');
            rowLabel.classList.add('axis-label');
            rowLabel.textContent = (0 - row * 0.25).toFixed(2);
            grid.appendChild(rowLabel);

            // 创建每行的其余单元格
            for (let col = 0; col < cols; col++) {
                const cell = document.createElement('input');
                cell.type = 'text';
                cell.classList.add('batch-cell');
                cell.dataset.spherical = (0 - row * 0.25).toFixed(2);
                cell.dataset.cylinder = (0 - col * 0.25).toFixed(2);
                grid.appendChild(cell);
            }
        }
    };
    createBatchGrid();

    // 监听鼠标按下事件，开始框选
    grid.addEventListener('mousedown', (e) => {
        const cell = e.target;
        if (cell.classList.contains('batch-cell')) {
            isDragging = true;
            startCell = cell;
            lastCell = cell;
            isRightClick = e.button === 2; // 判断是否为右键点击
            if (!isRightClick) {
                cell.value = '1'; // 左键单击时，设置值为 1
            }
            e.preventDefault(); // 禁止默认右键菜单
        }
    });

    // 监听鼠标移动事件，记录当前选中的区域
    grid.addEventListener('mousemove', (e) => {
        if (isDragging && startCell) {
            const cell = e.target;
            if (cell.classList.contains('batch-cell')) {
                lastCell = cell;
                if (isRightClick) {
                    cell.value = ''; // 右键拖拽时，清空值
                } else {
                    cell.value = '1'; // 左键拖拽时，设置值为 1
                }
                cell.classList.add('selected'); // 可选：在拖拽时添加视觉提示，显示选中的区域
            }
        }
    });

    // 监听鼠标松开事件，结束框选，并设置最后一个经过的单元格为当前选中单元格
    grid.addEventListener('mouseup', () => {
        if (isDragging) {
            isDragging = false;
            if (lastCell) {
                lastCell.focus();
            }
            startCell = null;
            lastCell = null;
        }
    });

    // 阻止右键菜单的默认行为（可选）
    document.addEventListener('contextmenu', (e) => {
        if (isDragging) {
            e.preventDefault(); // 禁止在拖拽过程中右键菜单弹出
        }
    });

    // 方向键快捷移动光标功能
    grid.addEventListener('keydown', (e) => {
        const cell = e.target;
        if (!cell.classList.contains('batch-cell')) return;

        let rowIndex = parseFloat(cell.dataset.spherical);
        let colIndex = parseFloat(cell.dataset.cylinder);
        let targetCell;

        switch (e.key) {
            case 'ArrowUp':
                targetCell = document.querySelector(`.batch-cell[data-spherical="${(rowIndex + 0.25).toFixed(2)}"][data-cylinder="${colIndex.toFixed(2)}"]`);
                break;
            case 'ArrowDown':
                targetCell = document.querySelector(`.batch-cell[data-spherical="${(rowIndex - 0.25).toFixed(2)}"][data-cylinder="${colIndex.toFixed(2)}"]`);
                break;
            case 'ArrowLeft':
                targetCell = document.querySelector(`.batch-cell[data-spherical="${rowIndex.toFixed(2)}"][data-cylinder="${(colIndex + 0.25).toFixed(2)}"]`);
                break;
            case 'ArrowRight':
                targetCell = document.querySelector(`.batch-cell[data-spherical="${rowIndex.toFixed(2)}"][data-cylinder="${(colIndex - 0.25).toFixed(2)}"]`);
                break;
            default:
                return; // 如果不是方向键，直接返回
        }

        if (targetCell) {
            targetCell.focus();
            targetCell.select();
            e.preventDefault();
        }
    });

    // 查看二维码图片的模态窗口功能
    viewQrcodesButton.addEventListener('click', () => {
        const modalBody = document.getElementById('qrcode-images');
        modalBody.innerHTML = ''; // 清空之前的内容

        const qrcodeImages = resultsDiv.querySelectorAll('img');
        console.log('Images found in resultsDiv:', qrcodeImages);

        qrcodeImages.forEach(img => {
            const clonedImg = img.cloneNode(true);
            modalBody.appendChild(clonedImg);
            console.log('Image added to modal:', clonedImg);
        });

        if (qrcodeImages.length === 0) {
            modalBody.innerHTML = '<p>没有生成的二维码图片。</p>';
        }

        const viewModal = new bootstrap.Modal(document.getElementById('view-modal'));
        viewModal.show();
    });

    // 打印所有二维码的功能
    printAllQrcodesButton.addEventListener('click', () => {
        const printWindow = window.open('', '_blank');
        const modalBody = document.getElementById('qrcode-images');
        printWindow.document.write('<html><head><title>打印二维码</title><style>body { margin: 0; padding: 0; } img { margin: 5px; }</style></head><body>');
        printWindow.document.write(modalBody.innerHTML);
        printWindow.document.write('</body></html>');
        printWindow.document.close();
        printWindow.print();

        printWindow.onafterprint = () => {
            const viewModal = bootstrap.Modal.getInstance(document.getElementById('view-modal'));
            viewModal.hide();
            modalBody.innerHTML = '';
        };
    });

    // 模态窗口关闭功能并清空单元格内容
    document.querySelectorAll('.btn-close, .btn-secondary').forEach(button => {
        button.addEventListener('click', (e) => {
            const modal = bootstrap.Modal.getInstance(e.target.closest('.modal'));
            modal.hide();

            // 清空所有单元格内容（不清空标签行和标签列）
            document.querySelectorAll('.batch-cell').forEach(cell => {
                cell.value = '';
                cell.classList.remove('selected');
            });
        });
    });
});
