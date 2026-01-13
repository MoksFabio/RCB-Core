document.addEventListener("DOMContentLoaded", () => {
    if (typeof ExcelJS === 'undefined') {
        console.error("Biblioteca ExcelJS (exceljs.min.js) não foi carregada. A geração de relatórios Excel não funcionará.");
    }

    const OuvidoriasApp = {
        elements: {
            loadingIndicator: document.getElementById('loading-indicator-ouv'),
            flashMessagesContainer: document.getElementById('flash-messages-container-ouv'),

            selectDemandLineBtn: document.getElementById('select-demand-line-assistant'),
            selectDemandMeanBtn: document.getElementById('select-demand-mean-assistant'),
            demandLineContainer: document.getElementById('demand-line-assistant-container'),
            demandMeanContainer: document.getElementById('demand-mean-assistant-container'),

            selectionScreen: document.getElementById('selection-screen'),
            mainAssistant: document.getElementById('main-assistant'),
            selectPermissionarias: document.getElementById('select-permissionarias'),
            selectConcessionarias: document.getElementById('select-concessionarias'),
            selectStppRmr: document.getElementById('select-stpp-rmr'),
            backToSelectionBtn: document.getElementById('back-to-selection-btn'),
            assistantTitle: document.getElementById('assistant-title'),
            validEmpresasInfo: document.getElementById('valid-empresas-info'),

            fileInputPermContainer: document.getElementById('file-input-perm-container'),
            fileInputConcContainer: document.getElementById('file-input-conc-container'),
            fileInputPerm: document.getElementById('file-input-perm'),
            fileInputConc: document.getElementById('file-input-conc'),
            dropZonePerm: document.getElementById('drop-zone-perm'),
            dropZoneConc: document.getElementById('drop-zone-conc'),

            codLinhas: document.getElementById('cod-linhas'),
            empresas: document.getElementById('empresas'),
            dataInicio: document.getElementById('data-inicio'),
            dataFim: document.getElementById('data-fim'),
            addPeriodBtn: document.getElementById('add-period-btn'),
            periodListContainer: document.getElementById('period-list-container'),
            removePeriodBtn: document.getElementById('remove-period-btn'),
            ouvidoriaId: document.getElementById('ouvidoria-id'),
            exportFormatLine: document.getElementById('export-format-line'),

            lineTipoDia: document.getElementById('line-tipo-dia') || document.createElement('select'),
            lineProcessEquivalentContainer: document.getElementById('line-process-equivalent-container'),
            lineProcessEquivalent: document.getElementById('line-process-equivalent') || document.createElement('input'),

            clearBtn: document.getElementById('clear-btn'),
            processBtn: document.getElementById('process-btn'),

            meanSelectionScreen: document.getElementById('mean-selection-screen'),
            mainMeanAssistant: document.getElementById('main-mean-assistant'),
            selectMeanPermissionarias: document.getElementById('select-mean-permissionarias'),
            selectMeanConcessionarias: document.getElementById('select-mean-concessionarias'),
            selectMeanStppRmr: document.getElementById('select-mean-stpp-rmr'),
            backToMeanSelectionBtn: document.getElementById('back-to-mean-selection-btn'),
            meanAssistantTitle: document.getElementById('mean-assistant-title'),
            meanValidEmpresasInfo: document.getElementById('mean-valid-empresas-info'),
            meanFileInputsContainer: document.getElementById('mean-file-inputs-container'),
            meanCodLinhas: document.getElementById('mean-cod-linhas'),
            meanEmpresas: document.getElementById('mean-empresas'),
            meanMultiMonthToggle: document.getElementById('mean-multi-month-toggle'),
            meanDateRangeContainer: document.getElementById('mean-date-range-container'),
            meanDataInicio: document.getElementById('mean-data-inicio'),
            meanDataFim: document.getElementById('mean-data-fim'),
            meanSpecificMonthsContainer: document.getElementById('mean-specific-months-container'),
            meanSpecificMonths: document.getElementById('mean-specific-months'),
            meanTipoDia: document.getElementById('mean-tipo-dia'),
            meanProcessEquivalentContainer: document.getElementById('mean-process-equivalent-container'),
            meanProcessEquivalent: document.getElementById('mean-process-equivalent'),
            meanOuvidoriaId: document.getElementById('mean-ouvidoria-id'),
            exportFormatMean: document.getElementById('export-format-mean'),
            meanClearBtn: document.getElementById('mean-clear-btn'),
            meanProcessBtn: document.getElementById('mean-process-btn'),
        },

        state: {
            lineProcessType: null,
            lineConfig: null,
            linePermFile: null,
            lineConcFile: null,
            lineLinesData: {},
            linePeriods: [],
            lineSelectedPeriodIndex: -1,

            meanProcessType: null,
            meanConfig: null,
            meanFiles: {},
            loadingInterval: null
        },

        config: {
            tiposDiaComuns: ["TODOS", "DUT", "SAB", "DOM"],

            lineConfigs: {
                permissionarias: {
                    title: "DEMANDA PERMISSIONÁRIAS",
                    empresas: ["BOA", "CAX", "CSR", "EME", "GLO", "SJT", "VML"],
                    colunasExcel: ['CDOPERADOR', 'CDLINHA', 'DTOPERACAO', 'NMEFETPASST', 'NMEFETPASS', 'DSDIATIPO'],
                    colunaPassageiros: 'NMEFETPASST',
                    colunaEquivalente: 'NMEFETPASS',
                    tituloExcel: "DEMANDA PERMISSIONÁRIAS"
                },
                concessionarias: {
                    title: "DEMANDA CONCESSIONÁRIAS",
                    empresas: ["CNO", "MOB"],
                    colunasExcel: ['CDOPERADOR', 'CDLINHA', 'DTOPERACAO', 'NMPASSTOTAL', 'DSDIATIPO', 'NMPASSEQUIVALENTE'],
                    colunaPassageiros: 'NMPASSTOTAL',
                    colunaEquivalente: 'NMPASSEQUIVALENTE',
                    tituloExcel: "DEMANDA CONCESSIONÁRIAS"
                },
                stpp_rmr: {
                    title: "DEMANDA STPP/RMR",
                    empresas: ["BOA", "CAX", "CSR", "EME", "GLO", "SJT", "VML", "CNO", "MOB"].sort(),
                    colunasExcel: ['CDOPERADOR', 'CDLINHA', 'DTOPERACAO', 'NMEFETPASST', 'NMEFETPASS', 'DSDIATIPO'],
                    colunaPassageiros: 'NMEFETPASST',
                    colunaEquivalente: 'NMEFETPASS',
                    tituloExcel: "DEMANDA STPP/RMR"
                }
            },
            meanConfigs: {
                permissionarias: {
                    title: "Cálculo de Média - Permissionárias",
                    empresas: ["BOA", "CAX", "CSR", "EME", "GLO", "SJT", "VML"],
                    tiposDia: ["TODOS", "DUT", "SAB", "DOM"],
                    fileConfigs: [
                        { id: 'permFile', label: 'Arquivo Permissionárias (.txt)', required: true, cols: ['CDOPERADOR', 'CDLINHA', 'DTOPERACAO', 'NMEFETPASST', 'DSDIATIPO'] }
                    ]
                },
                concessionarias: {
                    title: "Cálculo de Média - Concessionárias",
                    empresas: ["CNO", "MOB"],
                    tiposDia: ["TODOS", "DUT", "SAB", "DOM"],
                    fileConfigs: [
                        { id: 'concFile', label: 'Arquivo Concessionárias (.txt)', required: true, cols: ['CDOPERADOR', 'CDLINHA', 'DTOPERACAO', 'NMPASSTOTAL', 'NMPASSEQUIVALENTE'] },
                        { id: 'permLookupFile', label: 'Arquivo Permissionárias para consulta de Dia Útil (.txt)', required: false, cols: ['DTOPERACAO', 'DSDIATIPO'] }
                    ]
                },
                stpp_rmr: {
                    title: "Cálculo de Média - STPP/RMR",
                    empresas: ["BOA", "CAX", "CSR", "EME", "GLO", "SJT", "VML", "CNO", "MOB"].sort(),
                    tiposDia: ["TODOS", "DUT", "SAB", "DOM"],
                    fileConfigs: [
                        { id: 'permFile', label: 'Arquivo Permissionárias (.txt)', required: true, cols: ['CDOPERADOR', 'CDLINHA', 'DTOPERACAO', 'NMEFETPASST', 'DSDIATIPO'] },
                        { id: 'concFile', label: 'Arquivo Concessionárias (.txt)', required: true, cols: ['CDOPERADOR', 'CDLINHA', 'DTOPERACAO', 'NMPASSTOTAL', 'NMPASSEQUIVALENTE'] }
                    ]
                }
            }
        },

        init() {
            this.pageLogic.init();
        },
        pageLogic: {
            init() {
                OuvidoriasApp.pageLogic.processing.loadServerLinesFile();
                OuvidoriasApp.pageLogic.handlers.initLineDemandHandlers();
                OuvidoriasApp.pageLogic.handlers.initMeanDemandHandlers();
                OuvidoriasApp.pageLogic.handlers.initAssistantSwitcher();
            },

            utils: {
                getMonthName(idx) {
                    const months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
                    return months[idx - 1] || "";
                },

                showFlashMessage(message, type = 'info', duration = 7000) {
                    const container = OuvidoriasApp.elements.flashMessagesContainer;
                    if (!container) {
                        alert(message);
                        return;
                    }

                    const iconMap = { success: 'check_circle', error: 'warning', info: 'info' };
                    const icon = iconMap[type] || 'info';

                    const messageDiv = document.createElement('div');
                    messageDiv.className = `flash-message ${type}`;
                    messageDiv.setAttribute('role', 'alert');
                    messageDiv.style.transition = 'opacity 0.5s ease-out';
                    messageDiv.innerHTML = `
                        <span class="material-icons-outlined mr-3">${icon}</span>
                        <span>${message}</span>
                        <button type="button" class="ml-auto -mx-1.5 -my-1.5 p-1.5 inline-flex h-8 w-8 rounded-lg focus:ring-2 focus:ring-gray-300 dark:focus:ring-gray-600" aria-label="Dismiss">
                            <span class="material-icons-outlined text-sm current-color">close</span>
                        </button>
                    `;

                    messageDiv.querySelector('button').addEventListener('click', () => {
                        messageDiv.style.opacity = '0';
                        setTimeout(() => messageDiv.remove(), 500);
                    });

                    container.prepend(messageDiv);

                    if (duration > 0) {
                        setTimeout(() => {
                            if (messageDiv.parentElement) {
                                messageDiv.style.opacity = '0';
                                setTimeout(() => messageDiv.remove(), 500);
                            }
                        }, duration);
                    }
                    return messageDiv;
                },

                setLoading(isLoading, message = 'Gerando PDF...') {
                    const { processBtn, meanProcessBtn, clearBtn, meanClearBtn, backToSelectionBtn, backToMeanSelectionBtn } = OuvidoriasApp.elements;
                    // Buttons that should be just disabled
                    const buttonsToDisable = [clearBtn, meanClearBtn, backToSelectionBtn, backToMeanSelectionBtn];
                    // Buttons that will show the loading state
                    const actionButtons = [processBtn, meanProcessBtn];

                    if (isLoading) {
                        buttonsToDisable.forEach(btn => { if (btn) btn.disabled = true; });

                        actionButtons.forEach(btn => {
                            if (btn && !btn.disabled) { // Only affect active buttons (though usually there's only one visible)
                                btn.dataset.originalContent = btn.innerHTML;
                                btn.disabled = true;
                                btn.innerHTML = `<span class="material-icons-outlined animate-spin mr-2">autorenew</span> ${message}`;
                            } else if (btn) {
                                // If already disabled or other button, just ensure it stays disabled
                                btn.disabled = true;
                            }
                        });

                    } else {
                        buttonsToDisable.forEach(btn => { if (btn) btn.disabled = false; });

                        actionButtons.forEach(btn => {
                            if (btn) {
                                btn.disabled = false;
                                if (btn.dataset.originalContent) {
                                    btn.innerHTML = btn.dataset.originalContent;
                                } else {
                                    btn.innerHTML = `<span class="material-icons-outlined">play_arrow</span> Processar e Gerar Relatório`;
                                }
                            }
                        });
                    }
                },

                async loadImageBase64(url, maxWidth = null, maxHeight = null) {
                    return new Promise((resolve, reject) => {
                        const img = new Image();
                        img.crossOrigin = 'Anonymous';
                        img.onload = () => {
                            let width = img.width;
                            let height = img.height;

                            if (maxWidth && width > maxWidth) {
                                height *= maxWidth / width;
                                width = maxWidth;
                            }
                            if (maxHeight && height > maxHeight) {
                                width *= maxHeight / height;
                                height = maxHeight;
                            }

                            const canvas = document.createElement('canvas');
                            canvas.width = width;
                            canvas.height = height;
                            const ctx = canvas.getContext('2d');
                            ctx.drawImage(img, 0, 0, width, height);
                            // Reduce quality slightly to compress
                            resolve(canvas.toDataURL('image/png', 0.8));
                        };
                        img.onerror = () => {
                            console.warn("Não foi possível carregar a imagem: " + url);
                            resolve(null);
                        };
                        img.src = url;
                    });
                },

                async saveWorkbookWithPicker(workbook, fileName) {
                    if (typeof ExcelJS === 'undefined') {
                        OuvidoriasApp.pageLogic.utils.showFlashMessage("Erro: Biblioteca ExcelJS não carregada.", "error", 0);
                        return false;
                    }
                    const buffer = await workbook.xlsx.writeBuffer();
                    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

                    if ('showSaveFilePicker' in window) {
                        try {
                            const fileHandle = await window.showSaveFilePicker({
                                suggestedName: fileName,
                                types: [
                                    {
                                        description: 'Excel Workbook',
                                        accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
                                    },
                                ],
                            });
                            const writableStream = await fileHandle.createWritable();
                            await writableStream.write(blob);
                            await writableStream.close();
                            return true;
                        } catch (error) {
                            if (error.name === 'AbortError') return false;
                        }
                    }

                    const link = document.createElement('a');
                    link.href = URL.createObjectURL(blob);
                    link.download = fileName;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    URL.revokeObjectURL(link.href);
                    return true;
                },

                drawSummaryMatrix(worksheet, data, styles) {
                    const startColIndex = 7; // Coluna 7 = G

                    // --- AJUSTE DE POSIÇÃO (Iniciando na Linha 7) ---
                    const titleRowIdx = 7;   // Título "RESUMO..." fica na linha 7
                    const headerRowIdx = 8;  // Cabeçalho da tabela (OPERADOR, JAN, FEV...) na linha 8
                    let currentRowIdx = 9;   // Os dados começam na linha 9
                    // ------------------------------------------------

                    const matrixMap = {};
                    const uniqueMonths = new Set();
                    const uniqueOperators = new Set();
                    const monthTotals = {};

                    // ... (Lógica de processamento de dados permanece igual) ...
                    data.forEach(row => {
                        let dateObj = row.date;
                        if (!dateObj && row.DTOPERACAO) {
                            const match = row.DTOPERACAO.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
                            if (match) dateObj = new Date(match[3], match[2] - 1, match[1]);
                        }

                        if (!dateObj && row.DTOPERACAO_OBJ) dateObj = row.DTOPERACAO_OBJ;

                        if (!dateObj) return;

                        const monthKey = `${dateObj.getFullYear()}${String(dateObj.getMonth() + 1).padStart(2, '0')}`;
                        const op = row.CDOPERADOR;
                        const val = Number(row.VALOR_EXIBIDO || row.MÉDIA_EQUIVALENTE || row.MÉDIA_PASSAGEIROS || 0);

                        uniqueMonths.add(monthKey);
                        uniqueOperators.add(op);

                        if (!matrixMap[op]) matrixMap[op] = {};
                        if (!matrixMap[op][monthKey]) matrixMap[op][monthKey] = 0;
                        matrixMap[op][monthKey] += val;

                        if (!monthTotals[monthKey]) monthTotals[monthKey] = 0;
                        monthTotals[monthKey] += val;
                    });

                    const sortedMonths = Array.from(uniqueMonths).sort();
                    const sortedOperators = Array.from(uniqueOperators).sort();

                    const getShortMonthName = (mKey) => {
                        const m = parseInt(mKey.substring(4));
                        const months = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
                        return months[m - 1] || "";
                    };

                    // --- RENDERIZAÇÃO DO TÍTULO (LINHA 6) ---
                    const titleCell = worksheet.getCell(titleRowIdx, startColIndex);
                    titleCell.value = "RESUMO GERAL POR EMPRESA";
                    titleCell.style = styles.title;
                    titleCell.font = { name: 'Helvetica', size: 14, bold: true, color: { argb: 'FFD95F02' } };
                    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };

                    // Merge na linha do título (titleRowIdx)
                    worksheet.mergeCells(titleRowIdx, startColIndex, titleRowIdx, startColIndex + sortedMonths.length + 1);

                    // --- RENDERIZAÇÃO DO CABEÇALHO (LINHA 7) ---
                    worksheet.getCell(headerRowIdx, startColIndex).value = "OPERADOR";
                    worksheet.getCell(headerRowIdx, startColIndex).style = styles.matrixHeader;

                    sortedMonths.forEach((mKey, idx) => {
                        const y = mKey.substring(0, 4);
                        const label = `${getShortMonthName(mKey)}/${y}`;
                        const cell = worksheet.getCell(headerRowIdx, startColIndex + 1 + idx);
                        cell.value = label;
                        cell.style = styles.matrixHeader;
                        worksheet.getColumn(startColIndex + 1 + idx).width = 15;
                    });

                    const totalHeaderCell = worksheet.getCell(headerRowIdx, startColIndex + 1 + sortedMonths.length);
                    totalHeaderCell.value = "TOTAL GERAL";
                    totalHeaderCell.style = styles.matrixHeader;
                    worksheet.getColumn(startColIndex + 1 + sortedMonths.length).width = 18;
                    worksheet.getColumn(startColIndex).width = 15;

                    // --- RENDERIZAÇÃO DOS DADOS (LINHA 8 EM DIANTE) ---
                    let grandTotalAll = 0;

                    sortedOperators.forEach(op => {
                        worksheet.getCell(currentRowIdx, startColIndex).value = op;
                        worksheet.getCell(currentRowIdx, startColIndex).style = styles.operatorCell;

                        let rowTotal = 0;
                        sortedMonths.forEach((mKey, idx) => {
                            const val = matrixMap[op][mKey] || 0;
                            rowTotal += val;
                            const cell = worksheet.getCell(currentRowIdx, startColIndex + 1 + idx);
                            cell.value = val;
                            cell.style = styles.valueCell;
                            cell.numFmt = '#,##0.0';
                        });

                        const rowTotalCell = worksheet.getCell(currentRowIdx, startColIndex + 1 + sortedMonths.length);
                        rowTotalCell.value = rowTotal;
                        rowTotalCell.style = styles.valueCell;
                        rowTotalCell.font = { bold: true };
                        rowTotalCell.numFmt = '#,##0.0';

                        grandTotalAll += rowTotal;
                        currentRowIdx++;
                    });

                    // --- RODAPÉ (TOTAIS) ---
                    const footerRowIdx = currentRowIdx;
                    worksheet.getCell(footerRowIdx, startColIndex).value = "TOTAL GERAL";
                    worksheet.getCell(footerRowIdx, startColIndex).style = styles.totalRowCell;

                    sortedMonths.forEach((mKey, idx) => {
                        const val = monthTotals[mKey] || 0;
                        const cell = worksheet.getCell(footerRowIdx, startColIndex + 1 + idx);
                        cell.value = val;
                        cell.style = styles.totalRowCell;
                        cell.numFmt = '#,##0.0';
                    });

                    const grandTotalCell = worksheet.getCell(footerRowIdx, startColIndex + 1 + sortedMonths.length);
                    grandTotalCell.value = grandTotalAll;
                    grandTotalCell.style = styles.grandTotalCell;
                    grandTotalCell.numFmt = '#,##0.0';
                },

                async createExcelWorkbook(data, periods, fileName, logos = {}, extraInfo = {}) {
                    // Check if ExcelJS is loaded
                    if (!window.ExcelJS) {
                        OuvidoriasApp.pageLogic.utils.showFlashMessage("Erro Técnico: Biblioteca ExcelJS não carregada.", "error", 0);
                        return;
                    }

                    const workbook = new ExcelJS.Workbook();
                    const baseTitle = OuvidoriasApp.state.lineConfig.tituloExcel;
                    const todayStr = new Date().toLocaleDateString('pt-BR') + ' ' + new Date().toLocaleTimeString('pt-BR');

                    const colorPrimary = 'FFD95F02';
                    const colorSecondary = 'FF0038A8';
                    const colorWhite = 'FFFFFFFF';
                    const colorLightGray = 'FFF0F2F5';
                    const colorTotalBg = 'FFFFF0E6';
                    // const colorInput = 'FFFFFFCC'; // Unused
                    const colorBorder = 'FFBFBFBF';

                    const styles = {
                        title: {
                            font: { name: 'Helvetica', size: 18, bold: true, color: { argb: colorSecondary } },
                            alignment: { horizontal: 'right', vertical: 'middle' }
                        },
                        subtitle: {
                            font: { name: 'Helvetica', size: 12, color: { argb: 'FF505050' } },
                            alignment: { horizontal: 'right', vertical: 'middle' }
                        },
                        matrixHeader: {
                            font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorWhite } },
                            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorSecondary } },
                            alignment: { horizontal: 'center', vertical: 'middle' },
                            border: { top: { style: 'thin', color: { argb: colorWhite } }, left: { style: 'thin', color: { argb: colorWhite } }, bottom: { style: 'thin', color: { argb: colorWhite } }, right: { style: 'thin', color: { argb: colorWhite } } }
                        },
                        operatorCell: {
                            font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorSecondary } },
                            alignment: { horizontal: 'center', vertical: 'middle' },
                            border: { top: { style: 'thin', color: { argb: colorBorder } }, left: { style: 'thin', color: { argb: colorBorder } }, bottom: { style: 'thin', color: { argb: colorBorder } }, right: { style: 'thin', color: { argb: colorBorder } } }
                        },
                        valueCell: {
                            font: { name: 'Calibri', size: 11, color: { argb: 'FF555555' } },
                            alignment: { horizontal: 'center', vertical: 'middle' },
                            border: { top: { style: 'thin', color: { argb: colorBorder } }, left: { style: 'thin', color: { argb: colorBorder } }, bottom: { style: 'thin', color: { argb: colorBorder } }, right: { style: 'thin', color: { argb: colorBorder } } }
                        },
                        totalRowCell: {
                            font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorPrimary } },
                            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorTotalBg } },
                            alignment: { horizontal: 'center', vertical: 'middle' },
                            border: { top: { style: 'thin', color: { argb: colorPrimary } }, bottom: { style: 'thin', color: { argb: colorPrimary } }, left: { style: 'thin', color: { argb: colorBorder } }, right: { style: 'thin', color: { argb: colorBorder } } }
                        },
                        grandTotalCell: {
                            font: { name: 'Calibri', size: 12, bold: true, color: { argb: colorPrimary } },
                            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorTotalBg } },
                            alignment: { horizontal: 'center', vertical: 'middle' },
                            border: { top: { style: 'medium', color: { argb: colorPrimary } }, bottom: { style: 'medium', color: { argb: colorPrimary } }, left: { style: 'medium', color: { argb: colorPrimary } }, right: { style: 'medium', color: { argb: colorPrimary } } }
                        },
                        header: { font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorWhite } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorSecondary } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } } },
                        sectionHeader: { font: { name: 'Calibri', size: 12, bold: true, color: { argb: colorPrimary } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorLightGray } }, alignment: { horizontal: 'left', vertical: 'middle' } },
                        sectionHeaderRight: { font: { name: 'Calibri', size: 12, bold: true, color: { argb: colorPrimary } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorLightGray } }, alignment: { horizontal: 'right', vertical: 'middle' } },
                        cellLineTotal: { font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorPrimary } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorTotalBg } }, alignment: { horizontal: 'center', vertical: 'middle' } },
                        cellCompanyTotal: { font: { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } }, alignment: { horizontal: 'center', vertical: 'middle' } },
                        cellGrandTotalSection: { font: { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorSecondary } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'medium' }, left: { style: 'medium' }, bottom: { style: 'medium' }, right: { style: 'medium' } } },
                        cell: { font: { name: 'Calibri', size: 11, color: { argb: 'FF000000' } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin', left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } } } }
                    };
                    
                    // Pre-load logo once
                    let ctmLogoId = null;
                    if (logos.ctm) {
                        ctmLogoId = workbook.addImage({ base64: logos.ctm.split(',')[1], extension: 'png' });
                    }
                    
                    const totalPeriods = periods.length;

                    for (let pIdx = 0; pIdx < totalPeriods; pIdx++) {
                        const periodo = periods[pIdx];
                        // Update progress
                        const percent = Math.round(((pIdx) / totalPeriods) * 100);
                        OuvidoriasApp.pageLogic.utils.setLoading(true, `Gerando Excel... ${percent}%`);
                        
                        // Yield to UI
                        await new Promise(r => setTimeout(r, 0));

                        const dtInicioStr = periodo.inicio.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
                        const dtFimStr = periodo.fim.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
                        const sheetName = `${dtInicioStr.replace(/\//g, '')}_${dtFimStr.replace(/\//g, '')}`.substring(0, 30);

                        const worksheet = workbook.addWorksheet(sheetName, {
                            views: [{ state: 'frozen', ySplit: 6, showGridLines: false }]
                        });

                        // 1. Logo (Aumentando mesclagem vertical para linha 5)
                        worksheet.mergeCells(1, 1, 5, 1);
                        if (ctmLogoId !== null) {
                            worksheet.addImage(ctmLogoId, {
                                tl: { col: 0.2, row: 0.3 },
                                ext: { width: 280, height: 95 },
                                editAs: 'absolute'
                            });
                        }

                        // 2. Cabeçalho (Linhas 1 a 5)

                        // LINHA 1: Título Principal (Alinhado à Direita)
                        worksheet.mergeCells(1, 4, 1, 13); // Coluna D(4) até M(13)
                        worksheet.getCell('D1').value = "GRANDE RECIFE CONSÓRCIO DE TRANSPORTE";
                        worksheet.getCell('D1').style = styles.title;
                        worksheet.getCell('D1').font = { ...styles.title.font, size: 14, color: { argb: 'FF808080' } };

                        // LINHA 2: Título do Relatório (Alinhado à Direita)
                        worksheet.mergeCells(2, 4, 2, 13);
                        worksheet.getCell('D2').value = baseTitle;
                        worksheet.getCell('D2').style = styles.title;

                        // LINHA 3: DGFC (NOVO - Alinhado à Direita, abaixo do título)
                        worksheet.mergeCells(3, 4, 3, 13);
                        worksheet.getCell('D3').value = "DGFC - DIVISÃO DE GESTÃO FINANCEIRA DOS CONTRATOS";
                        worksheet.getCell('D3').style = styles.subtitle; // Usa o estilo subtitle que já é alinhado à direita

                        // LINHA 4: Período (Alinhado à ESQUERDA)
                        worksheet.mergeCells(4, 4, 4, 13);
                        let subtitleText = `PERÍODO: ${dtInicioStr} A ${dtFimStr}`;
                        if (extraInfo.tipoDia) subtitleText += ` | TIPO DIA: ${extraInfo.tipoDia}`;
                        if (extraInfo.equivalente) subtitleText += ` | INCLUI PASS. EQUIVALENTE`;
                        worksheet.getCell('D4').value = subtitleText;
                        worksheet.getCell('D4').style = styles.subtitle;
                        worksheet.getCell('D4').alignment = { horizontal: 'left', vertical: 'middle' };

                        // LINHA 5: Gerado em (Alinhado à ESQUERDA)
                        worksheet.mergeCells(5, 4, 5, 13);
                        worksheet.getCell('D5').value = `Gerado em: ${todayStr}`;
                        worksheet.getCell('D5').style = styles.subtitle;
                        worksheet.getCell('D5').font = { size: 9, italic: true, color: { argb: 'FF808080' } };
                        worksheet.getCell('D5').alignment = { horizontal: 'left', vertical: 'middle' };

                        worksheet.addRow([]);

                        const sheetDataFiltered = data.filter(row => row.DTOPERACAO_NUM >= periodo.inicioNum && row.DTOPERACAO_NUM <= periodo.fimNum);

                        // *** AQUI ESTAVA O ERRO: A CHAMADA drawSummaryMatrix FOI REMOVIDA DAQUI ***

                        const linesMap = {};
                        sheetDataFiltered.forEach(row => {
                            const formattedDate = row.DTOPERACAO;
                            const passVal = Number(row.VALOR_EXIBIDO || 0);
                            const opTxt = row.CDOPERADOR;
                            const codeTxt = parseInt(row.CDLINHA).toString();
                            const lineKey = `${opTxt}_${codeTxt}`;
                            const lineName = OuvidoriasApp.state.lineLinesData[lineKey] || 'NOME NÃO ENCONTRADO';

                            if (!linesMap[lineKey]) {
                                linesMap[lineKey] = { operador: opTxt, linha: codeTxt, nome: lineName, rows: [] };
                            }
                            linesMap[lineKey].rows.push({ ...row, PASSAGEIROS: passVal, DTOPERACAO: formattedDate });
                        });

                        const sortedLineKeys = Object.keys(linesMap).sort((a, b) => {
                            const grpA = linesMap[a];
                            const grpB = linesMap[b];
                            if (grpA.operador !== grpB.operador) return grpA.operador.localeCompare(grpB.operador);
                            return parseInt(grpA.linha) - parseInt(grpB.linha);
                        });

                        let currentOperator = null;
                        let operatorPeriodTotal = 0;
                        let periodGrandTotal = 0;

                        const totalTasks = sortedLineKeys.length;
                        
                        // Process lines in chunks to allow UI update if many lines
                        for (let i = 0; i < totalTasks; i++) {
                            // Yield every N lines within the sheet loop as well if needed
                            if (i % 50 === 0 && i > 0) {
                                await new Promise(r => setTimeout(r, 0));
                            }

                            const lineKey = sortedLineKeys[i];
                            const lineGroup = linesMap[lineKey];

                            if (currentOperator !== lineGroup.operador) {
                                if (currentOperator !== null) {
                                    worksheet.addRow([]);
                                    const opRow = worksheet.addRow([`TOTAL ${currentOperator}`, '', '', '', operatorPeriodTotal]);
                                    worksheet.mergeCells(opRow.number, 1, opRow.number, 4);
                                    opRow.eachCell((cell, col) => { cell.style = styles.cellCompanyTotal; if (col === 5) cell.numFmt = '#,##0'; });
                                    worksheet.addRow([]); worksheet.addRow([]);
                                }
                                currentOperator = lineGroup.operador;
                                operatorPeriodTotal = 0;
                            }

                            const monthsMap = {};
                            lineGroup.rows.forEach(row => {
                                const [d, m, y] = row.DTOPERACAO.split('/');
                                const mKey = `${y}${m}`;
                                if (!monthsMap[mKey]) monthsMap[mKey] = { idx: parseInt(m), yr: parseInt(y), rows: [] };
                                monthsMap[mKey].rows.push(row);
                            });

                            const sortedMonthKeys = Object.keys(monthsMap).sort((a, b) => parseInt(a) - parseInt(b));

                            for (const mKey of sortedMonthKeys) {
                                const mData = monthsMap[mKey];
                                const mName = `${OuvidoriasApp.pageLogic.utils.getMonthName(mData.idx)} ${mData.yr}`;
                                const q1 = [], q2 = [];
                                mData.rows.forEach(r => {
                                    const d = parseInt(r.DTOPERACAO.split('/')[0]);
                                    if (d <= 15) q1.push(r); else q2.push(r);
                                });

                                const maxRows = Math.max(q1.length, q2.length);
                                let lineMonthTotal = 0;
                                const combinedRows = [];

                                for (let k = 0; k < maxRows; k++) {
                                    const r1 = q1[k] || {};
                                    const r2 = q2[k] || {};
                                    const op = r1.CDOPERADOR || r2.CDOPERADOR || lineGroup.operador;
                                    const p1 = r1.PASSAGEIROS ?? null;
                                    const p2 = r2.PASSAGEIROS ?? null;
                                    if (p1) lineMonthTotal += p1;
                                    if (p2) lineMonthTotal += p2;
                                    combinedRows.push({ operador: op, data1: r1.DTOPERACAO || '', pass1: p1, data2: r2.DTOPERACAO || '', pass2: p2 });
                                }
                                operatorPeriodTotal += lineMonthTotal;
                                periodGrandTotal += lineMonthTotal;

                                const sRow = worksheet.addRow([`LINHA ${lineGroup.linha} – ${lineGroup.nome}`, '', '', '', mName.toUpperCase()]);
                                const sKey = `${lineGroup.operador}_${lineGroup.linha}`;
                                worksheet.mergeCells(sRow.number, 1, sRow.number, 4);
                                sRow.getCell(1).style = styles.sectionHeader;
                                sRow.getCell(5).style = styles.sectionHeaderRight;

                                const ph = extraInfo.equivalente ? 'PASS. EQUIV.' : 'PASSAGEIROS';
                                const thRow = worksheet.addRow(['OPERADOR', '1ª QUINZENA', ph, '2ª QUINZENA', ph]);
                                thRow.eachCell((c, i) => { if (i <= 5) c.style = styles.header; });

                                combinedRows.forEach(cRow => {
                                    const row = worksheet.addRow([cRow.operador, cRow.data1, cRow.pass1, cRow.data2, cRow.pass2]);
                                    for (let c = 1; c <= 5; c++) {
                                        const cell = row.getCell(c);
                                        cell.style = styles.cell;
                                        if ((c === 3 || c === 5) && cell.value !== null) cell.numFmt = '#,##0';
                                    }
                                });

                                const tRow = worksheet.addRow(['TOTAL LINHA', '', '', '', lineMonthTotal]);
                                worksheet.mergeCells(tRow.number, 2, tRow.number, 4);
                                tRow.getCell(1).style = styles.cellLineTotal;
                                tRow.getCell(2).style = styles.cellLineTotal;
                                tRow.getCell(5).style = styles.cellLineTotal;
                                tRow.getCell(5).numFmt = '#,##0';
                                worksheet.addRow([]);
                            }
                        }

                        if (currentOperator !== null) {
                            worksheet.addRow([]);
                            const opRow = worksheet.addRow([`TOTAL ${currentOperator}`, '', '', '', operatorPeriodTotal]);
                            worksheet.mergeCells(opRow.number, 1, opRow.number, 4);
                            opRow.eachCell((cell, col) => { cell.style = styles.cellCompanyTotal; if (col === 5) cell.numFmt = '#,##0'; });
                            worksheet.addRow([]);
                        }

                        worksheet.addRow([]);
                        const gRow = worksheet.addRow(['TOTAL GERAL DO PERÍODO', '', '', '', periodGrandTotal]);
                        worksheet.mergeCells(gRow.number, 1, gRow.number, 4);
                        gRow.eachCell((c, col) => { c.style = styles.cellGrandTotalSection; if (col === 5) c.numFmt = '#,##0'; });

                        // *** AQUI ESTÁ A CORREÇÃO: CHAMADA MOVIDA PARA O FINAL DO LOOP ***
                        OuvidoriasApp.pageLogic.utils.drawSummaryMatrix(worksheet, sheetDataFiltered, styles);
                        // ***************************************************************

                        worksheet.getColumn(1).width = 15;
                        worksheet.getColumn(2).width = 20;
                        worksheet.getColumn(3).width = 15;
                        worksheet.getColumn(4).width = 20;
                        worksheet.getColumn(5).width = 15;
                    }

                    if (workbook.worksheets.length === 0) {
                        workbook.addWorksheet("Sem Dados").addRow(["Nenhum período selecionado ou dados encontrados."]);
                    }

                    OuvidoriasApp.pageLogic.utils.setLoading(true, `Finalizando arquivo... 100%`);
                    await new Promise(r => setTimeout(r, 200));

                    await OuvidoriasApp.pageLogic.utils.saveWorkbookWithPicker(workbook, fileName);
                },

                addMeanWorksheetToWorkbook(workbook, sheetName, data, inputs, periodStr, logos) {
                    const baseTitle = OuvidoriasApp.state.meanConfig.title;
                    const todayStr = new Date().toLocaleDateString('pt-BR') + ' ' + new Date().toLocaleTimeString('pt-BR');

                    const colorPrimary = 'FFD95F02';
                    const colorSecondary = 'FF0038A8';
                    const colorWhite = 'FFFFFFFF';
                    const colorLightGray = 'FFF0F2F5';
                    const colorTotalBg = 'FFFFF0E6';
                    const colorBorder = 'FFBFBFBF';

                    const styles = {
                        title: { font: { name: 'Helvetica', size: 18, bold: true, color: { argb: colorSecondary } }, alignment: { horizontal: 'left', vertical: 'middle' } },
                        subtitle: { font: { name: 'Helvetica', size: 12, color: { argb: 'FF505050' } }, alignment: { horizontal: 'right', vertical: 'middle' } },
                        header: { font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorWhite } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorSecondary } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } } },
                        cell: { font: { name: 'Calibri', size: 11, color: { argb: 'FF000000' } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } } },
                        matrixHeader: { font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorWhite } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorSecondary } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin', color: { argb: colorWhite } }, left: { style: 'thin', color: { argb: colorWhite } }, bottom: { style: 'thin', color: { argb: colorWhite } }, right: { style: 'thin', color: { argb: colorWhite } } } },
                        operatorCell: { font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorSecondary } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin', color: { argb: colorBorder } }, left: { style: 'thin', color: { argb: colorBorder } }, bottom: { style: 'thin', color: { argb: colorBorder } }, right: { style: 'thin', color: { argb: colorBorder } } } },
                        valueCell: { font: { name: 'Calibri', size: 11, color: { argb: 'FF555555' } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin', color: { argb: colorBorder } }, left: { style: 'thin', color: { argb: colorBorder } }, bottom: { style: 'thin', color: { argb: colorBorder } }, right: { style: 'thin', color: { argb: colorBorder } } } },
                        totalRowCell: { font: { name: 'Calibri', size: 11, bold: true, color: { argb: colorPrimary } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorTotalBg } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin', color: { argb: colorPrimary } }, bottom: { style: 'thin', color: { argb: colorPrimary } }, left: { style: 'thin', color: { argb: colorBorder } }, right: { style: 'thin', color: { argb: colorBorder } } } },
                        grandTotalCell: { font: { name: 'Calibri', size: 12, bold: true, color: { argb: colorPrimary } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: colorTotalBg } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'medium', color: { argb: colorPrimary } }, bottom: { style: 'medium', color: { argb: colorPrimary } }, left: { style: 'medium', color: { argb: colorPrimary } }, right: { style: 'medium', color: { argb: colorPrimary } } } }
                    };

                    const worksheet = workbook.addWorksheet(sheetName.substring(0, 30), {
                        views: [{ state: 'frozen', ySplit: 7, showGridLines: false }]
                    });

                    // Logo (Mescla até linha 5)
                    worksheet.mergeCells(1, 1, 5, 1);
                    if (logos.ctm) {
                        const ctmLogoId = workbook.addImage({ base64: logos.ctm.split(',')[1], extension: 'png' });
                        worksheet.addImage(ctmLogoId, {
                            tl: { col: 0.2, row: 0.3 },
                            ext: { width: 280, height: 95 },
                            editAs: 'absolute'
                        });
                    }

                    // LINHA 1
                    worksheet.mergeCells(1, 4, 1, 13);
                    worksheet.getCell('D1').value = "GRANDE RECIFE CONSÓRCIO DE TRANSPORTE";
                    worksheet.getCell('D1').style = styles.title;
                    worksheet.getCell('D1').font = { ...styles.title.font, size: 14, color: { argb: 'FF808080' } };

                    // LINHA 2
                    worksheet.mergeCells(2, 4, 2, 13);
                    worksheet.getCell('D2').value = baseTitle;
                    worksheet.getCell('D2').style = styles.title;

                    // LINHA 3 (DGFC)
                    worksheet.mergeCells(3, 4, 3, 13);
                    worksheet.getCell('D3').value = "DGFC - DIVISÃO DE GESTÃO FINANCEIRA DOS CONTRATOS";
                    worksheet.getCell('D3').style = styles.subtitle;

                    // LINHA 4 (Referência - Esquerda)
                    worksheet.mergeCells(4, 4, 4, 13);
                    let subtitleText = `REFERÊNCIA: ${periodStr} | TIPO DIA: ${inputs.tipoDia}`;
                    if (inputs.processEquivalent) subtitleText += ` | INCLUI PASS. EQUIVALENTE`;
                    worksheet.getCell('D4').value = subtitleText;
                    worksheet.getCell('D4').style = styles.subtitle;
                    worksheet.getCell('D4').alignment = { horizontal: 'left', vertical: 'middle' };

                    // LINHA 5 (Gerado em - Esquerda)
                    worksheet.mergeCells(5, 4, 5, 13);
                    worksheet.getCell('D5').value = `Gerado em: ${todayStr}`;
                    worksheet.getCell('D5').style = styles.subtitle;
                    worksheet.getCell('D5').font = { size: 9, italic: true, color: { argb: 'FF808080' } };
                    worksheet.getCell('D5').alignment = { horizontal: 'left', vertical: 'middle' };

                    // Row 6 (Empty)
                    worksheet.addRow([]);

                    // Removed Summary Matrix as per request
                    // The main table will now start on Row 7

                    const cols = ['OPERADOR', 'CÓD. LINHA', 'NOME DA LINHA', 'DIAS', 'MÉDIA PASS.'];
                    if (inputs.processEquivalent) cols.push('MÉDIA EQUIV.');

                    const headerRow = worksheet.addRow(cols);
                    headerRow.eachCell(cell => cell.style = styles.header);

                    data.forEach(row => {
                        const values = [
                            row.CDOPERADOR,
                            row.CDLINHA,
                            OuvidoriasApp.state.lineLinesData[`${row.CDOPERADOR}_${parseInt(row.CDLINHA)}`] || 'NOME NÃO ENCONTRADO',
                            row.DIAS_CONTABILIZADOS,
                            row.MÉDIA_PASSAGEIROS
                        ];
                        if (inputs.processEquivalent) values.push(row.MÉDIA_EQUIVALENTE);

                        const r = worksheet.addRow(values);
                        r.eachCell((cell, colNum) => {
                            cell.style = styles.cell;
                            // Columns 4 (DIAS), 5 (MEDIA), 6 (EQUIV) are numeric
                            if (colNum >= 4) cell.numFmt = '#,##0';
                        });
                    });

                    worksheet.getColumn(1).width = 15;
                    worksheet.getColumn(2).width = 12;
                    worksheet.getColumn(3).width = 35;
                    worksheet.getColumn(4).width = 10;
                    worksheet.getColumn(5).width = 18;
                    if (inputs.processEquivalent) worksheet.getColumn(6).width = 18;
                },

                async generatePDF(data, title, fileNameBase, type = 'line', infoExtra = "", logos = {}) {
                    if (!window.jspdf) {
                        OuvidoriasApp.pageLogic.utils.showFlashMessage("Erro: Biblioteca jsPDF não carregada.", "error", 0);
                        return false;
                    }

                    const { jsPDF } = window.jspdf;
                    const doc = new jsPDF({
                        orientation: 'l',
                        unit: 'mm',
                        format: 'a4',
                        compress: true
                    });

                    if (typeof doc.autoTable !== 'function') {
                        OuvidoriasApp.pageLogic.utils.showFlashMessage("Erro Técnico: Plugin 'AutoTable' não carregado.", "error", 0);
                        return false;
                    }

                    const ctmLogoBase64 = await OuvidoriasApp.pageLogic.utils.loadImageBase64('/static/images/ctm.png');

                    const primaryColor = [217, 95, 2];
                    const secondaryColor = [0, 56, 168];
                    const lightGray = [240, 242, 245];
                    const darkText = [44, 62, 80];

                    const headerHeight = 65;
                    const marginBottom = 18;

                    const drawHeader = (docData) => {
                        const currentDoc = docData.doc;
                        const pageSize = currentDoc.internal.pageSize;
                        const pageWidth = pageSize.width;

                        currentDoc.setFillColor(primaryColor[0], primaryColor[1], primaryColor[2]);
                        currentDoc.rect(0, 0, pageWidth, 5, 'F');

                        const logoY = 7;
                        let currentX = 14;

                        if (ctmLogoBase64) {
                            const logoRatio = 2.66;
                            const logoHeight = 42;
                            const logoWidth = logoHeight * logoRatio;
                            const offsetY = 8;

                            currentDoc.addImage(ctmLogoBase64, 'PNG', currentX, logoY + offsetY, logoWidth, logoHeight);
                            currentDoc.link(currentX, logoY + offsetY, logoWidth, logoHeight, { url: 'https://www.granderecife.pe.gov.br/' });
                        }

                        currentDoc.setFont("helvetica", "bold");
                        currentDoc.setTextColor(secondaryColor[0], secondaryColor[1], secondaryColor[2]);
                        currentDoc.setFontSize(18);
                        currentDoc.text("GRANDE RECIFE CONSÓRCIO DE TRANSPORTE", pageWidth - 14, 20, { align: 'right' });

                        currentDoc.setFontSize(14);
                        currentDoc.setTextColor(100, 100, 100);
                        currentDoc.text("DGFC - DIVISÃO DE GESTÃO FINANCEIRA DOS CONTRATOS", pageWidth - 14, 28, { align: 'right' });

                        currentDoc.setFontSize(22);
                        currentDoc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
                        currentDoc.text(title.toUpperCase(), pageWidth - 14, 40, { align: 'right' });

                        currentDoc.setFont("helvetica", "normal");
                        currentDoc.setFontSize(12);
                        currentDoc.setTextColor(80, 80, 80);
                        currentDoc.text(infoExtra || "Relatório Analítico", pageWidth - 14, 48, { align: 'right' });

                        currentDoc.setDrawColor(200, 200, 200);
                        currentDoc.setLineWidth(0.5);
                        currentDoc.line(14, 62, pageWidth - 14, 62);
                    };

                    const drawFooter = (docData) => {
                        const currentDoc = docData.doc;
                        const pageSize = currentDoc.internal.pageSize;
                        const pageWidth = pageSize.width;
                        const pageHeight = pageSize.height;

                        currentDoc.setDrawColor(primaryColor[0], primaryColor[1], primaryColor[2]);
                        currentDoc.setLineWidth(1);
                        currentDoc.line(14, pageHeight - 15, pageWidth - 14, pageHeight - 15);

                        const today = new Date();
                        const dateStr = today.toLocaleDateString('pt-BR') + ' ' + today.toLocaleTimeString('pt-BR');

                        currentDoc.setFontSize(9);
                        currentDoc.setTextColor(100, 100, 100);
                        currentDoc.text(`Gerado em: ${dateStr}`, 14, pageHeight - 8);

                        currentDoc.text("Grande Recife Consórcio de Transporte - Relatório Oficial", pageWidth / 2, pageHeight - 8, { align: 'center' });

                        const pageNumber = `Pág. ${currentDoc.internal.getNumberOfPages()}`;
                        currentDoc.text(pageNumber, pageWidth - 14, pageHeight - 8, { align: 'right' });
                    };

                    const formatData = (items) => {
                        return items.map(item => {
                            let formattedDate = item.DTOPERACAO;
                            if (type === 'line' && item.DTOPERACAO_OBJ) {
                                const d = item.DTOPERACAO_OBJ;
                                const day = String(d.getUTCDate()).padStart(2, '0');
                                const month = String(d.getUTCMonth() + 1).padStart(2, '0');
                                const year = d.getUTCFullYear();
                                formattedDate = `${day}/${month}/${year}`;
                            }

                            let passVal = item.PASSAGEIROS;
                            if (type === 'line') {
                                passVal = (item.VALOR_EXIBIDO || 0).toLocaleString('pt-BR');
                            }

                            const opTxt = item.CDOPERADOR;
                            const codeTxt = parseInt(item.CDLINHA).toString();
                            const lineKey = `${opTxt}_${codeTxt}`;

                            const lineName = OuvidoriasApp.state.lineLinesData[lineKey] || 'NOME NÃO ENCONTRADO';

                            return {
                                CDOPERADOR: item.CDOPERADOR,
                                CDLINHA: item.CDLINHA,
                                LINHA_NOME: lineName,
                                DTOPERACAO: formattedDate,
                                PASSAGEIROS: passVal,
                                DIAS_CONTABILIZADOS: item.DIAS_CONTABILIZADOS,
                                MÉDIA_PASSAGEIROS: item.MÉDIA_PASSAGEIROS ? Math.round(item.MÉDIA_PASSAGEIROS).toLocaleString('pt-BR') : '',
                                MÉDIA_EQUIVALENTE: item.MÉDIA_EQUIVALENTE ? Math.round(item.MÉDIA_EQUIVALENTE).toLocaleString('pt-BR') : '-'
                            };
                        });
                    };

                    OuvidoriasApp.pageLogic.utils.setLoading(true, 'Preparando dados para PDF...');
                    await new Promise(r => setTimeout(r, 10));

                    const formattedBody = formatData(data);
                    let startY = headerHeight + 5;
                    let firstPageRendered = false;

                    const getMonthName = (idx) => {
                        const months = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
                        return months[idx - 1] || "";
                    };

                    if (type === 'line') {
                        const linesMap = {};
                        const grandTotalMap = {};

                        formattedBody.forEach(row => {
                            const key = `${row.CDOPERADOR}|${row.CDLINHA}`;
                            if (!linesMap[key]) {
                                linesMap[key] = {
                                    operador: row.CDOPERADOR,
                                    linha: row.CDLINHA,
                                    nome: row.LINHA_NOME,
                                    rows: []
                                };
                            }
                            linesMap[key].rows.push(row);
                        });

                        const sortedLineKeys = Object.keys(linesMap).sort((a, b) => {
                            const grpA = linesMap[a];
                            const grpB = linesMap[b];
                            if (grpA.operador !== grpB.operador) return grpA.operador.localeCompare(grpB.operador);
                            return parseInt(grpA.linha) - parseInt(grpB.linha);
                        });

                        const pdfTasks = [];

                        for (const lineKey of sortedLineKeys) {
                            const lineGroup = linesMap[lineKey];
                            const monthsMap = {};

                            lineGroup.rows.forEach(row => {
                                const [day, month, year] = row.DTOPERACAO.split('/');
                                const monthKey = `${year}${month}`;
                                if (!monthsMap[monthKey]) {
                                    monthsMap[monthKey] = {
                                        monthIndex: parseInt(month, 10),
                                        year: parseInt(year, 10),
                                        rows: []
                                    };
                                }
                                monthsMap[monthKey].rows.push(row);
                            });

                            const sortedMonthKeys = Object.keys(monthsMap).sort((a, b) => parseInt(a) - parseInt(b));

                            for (const monthKey of sortedMonthKeys) {
                                pdfTasks.push({
                                    type: 'line_month_table',
                                    lineGroup: lineGroup,
                                    monthData: monthsMap[monthKey]
                                });

                                const monthData = monthsMap[monthKey];
                                const sortableMonthKey = `${monthData.year}-${String(monthData.monthIndex).padStart(2, '0')}`;
                                const op = lineGroup.operador;

                                let monthLineTotal = 0;
                                monthData.rows.forEach(r => {
                                    const val = parseFloat(String(r.PASSAGEIROS).replace(/\./g, '').replace(',', '.')) || 0;
                                    monthLineTotal += val;
                                });

                                if (!grandTotalMap[op]) grandTotalMap[op] = {};
                                if (!grandTotalMap[op][sortableMonthKey]) grandTotalMap[op][sortableMonthKey] = 0;
                                grandTotalMap[op][sortableMonthKey] += monthLineTotal;
                            }
                        }

                        if (Object.keys(grandTotalMap).length > 0) {
                            const uniqueMonths = new Set();
                            Object.values(grandTotalMap).forEach(mMap => Object.keys(mMap).forEach(k => uniqueMonths.add(k)));
                            const sortedMonths = Array.from(uniqueMonths).sort();

                            const summaryCols = [{ header: 'OPERADOR', dataKey: 'OPERADOR' }];
                            sortedMonths.forEach(mKey => {
                                const [y, m] = mKey.split('-');
                                const mName = getMonthName(parseInt(m)).substring(0, 3) + '/' + y;
                                summaryCols.push({ header: mName.toUpperCase(), dataKey: mKey });
                            });
                            summaryCols.push({ header: 'TOTAL GERAL', dataKey: 'TOTAL' });

                            const summaryRows = [];
                            const sortedOps = Object.keys(grandTotalMap).sort();

                            sortedOps.forEach(op => {
                                const row = { OPERADOR: op };
                                let lineTotal = 0;
                                sortedMonths.forEach(mKey => {
                                    const val = grandTotalMap[op][mKey] || 0;
                                    row[mKey] = val.toLocaleString('pt-BR');
                                    lineTotal += val;
                                });
                                row['TOTAL'] = lineTotal.toLocaleString('pt-BR');
                                summaryRows.push(row);
                            });

                            const totalGeralRow = { OPERADOR: 'TOTAL GERAL' };
                            let allTotal = 0;
                            sortedMonths.forEach(mKey => {
                                let colSum = 0;
                                sortedOps.forEach(op => colSum += (grandTotalMap[op][mKey] || 0));
                                totalGeralRow[mKey] = colSum.toLocaleString('pt-BR');
                                allTotal += colSum;
                            });
                            totalGeralRow['TOTAL'] = allTotal.toLocaleString('pt-BR');
                            summaryRows.push(totalGeralRow);

                            doc.setFontSize(14);
                            doc.setFont("helvetica", "bold");
                            doc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
                            doc.text("RESUMO GERAL POR EMPRESA", 14, headerHeight + 5);

                            doc.autoTable({
                                columns: summaryCols,
                                body: summaryRows,
                                startY: headerHeight + 10,
                                theme: 'grid',
                                styles: { fontSize: 8, textColor: darkText, cellPadding: 2, valign: 'middle', halign: 'center' },
                                headStyles: { fillColor: secondaryColor, textColor: [255, 255, 255], fontStyle: 'bold' },
                                columnStyles: { 0: { halign: 'center', fontStyle: 'bold', textColor: secondaryColor, width: 25 } },
                                didDrawPage: function (data) { drawHeader(data); drawFooter(data); },
                                didParseCell: function (data) {
                                    if (data.section === 'body' && data.row.index === summaryRows.length - 1) {
                                        data.cell.styles.fontStyle = 'bold';
                                        data.cell.styles.fillColor = [255, 240, 230];
                                        data.cell.styles.textColor = primaryColor;
                                    }
                                },
                                didDrawCell: function (data) {
                                    if (data.section === 'body' && data.column.index === 0) {
                                        const op = data.cell.raw;
                                        if (logos[op]) {
                                             const dim = data.cell.height - 2; // deixa uma margem de 1mm
                                             const textPos = data.cell.getTextPos();
                                             // Ajusta imagem pequena à esquerda do texto ou centralizada se não couber junto
                                             // Vamos tentar colocar à esquerda do texto se houver espaço, ou apenas desenhar a imagem
                                             // Como pedido "logos devem estar pequenas para caberem certinho"
                                             
                                            // Desenhar imagem centralizada verticalmente, alinhada à esquerda da célula com um padding
                                            // Aumentando para 10mm (ou menor se a altura da célula limitar)
                                            // Vamos garantir que não exceda a altura da célula com uma margem de 1mm
                                            const padding = 1;
                                            const maxH = data.cell.height - (padding * 2);
                                            const targetSize = 10;
                                            const imgSize = Math.min(targetSize, maxH > 0 ? maxH : targetSize);

                                            // Add image
                                            try {
                                                const x = data.cell.x + 2;
                                                const y = data.cell.y + (data.cell.height - imgSize) / 2;
                                                doc.addImage(logos[op], 'PNG', x, y, imgSize, imgSize);
                                            } catch (e) {
                                                console.warn("Erro ao desenhar logo no PDF para " + op);
                                            }
                                        }
                                    }
                                }
                            });

                            startY = doc.lastAutoTable.finalY + 15;
                            firstPageRendered = true;
                        }

                        const totalTasks = pdfTasks.length;
                        for (let i = 0; i < totalTasks; i++) {
                            if (i % 10 === 0) {
                                const percent = Math.round((i / totalTasks) * 100);
                                OuvidoriasApp.pageLogic.utils.setLoading(true, `Gerando PDF... ${percent}%`);
                                await new Promise(r => setTimeout(r, 0));
                            }

                            const task = pdfTasks[i];
                            const lineGroup = task.lineGroup;
                            const monthData = task.monthData;
                            const monthNameFull = `${getMonthName(monthData.monthIndex)} ${monthData.year}`;

                            const quinzena1 = [];
                            const quinzena2 = [];

                            monthData.rows.forEach(row => {
                                const day = parseInt(row.DTOPERACAO.split('/')[0], 10);
                                if (!isNaN(day) && day <= 15) {
                                    quinzena1.push(row);
                                } else {
                                    quinzena2.push(row);
                                }
                            });

                            let monthLineTotal = 0;
                            const combinedRows = [];
                            const maxRows = Math.max(quinzena1.length, quinzena2.length);

                            for (let k = 0; k < maxRows; k++) {
                                const r1 = quinzena1[k] || {};
                                const r2 = quinzena2[k] || {};
                                const baseOp = r1.CDOPERADOR || r2.CDOPERADOR || lineGroup.operador;

                                const p1Str = r1.PASSAGEIROS ? String(r1.PASSAGEIROS) : '0';
                                const p2Str = r2.PASSAGEIROS ? String(r2.PASSAGEIROS) : '0';
                                const p1Num = parseFloat(p1Str.replace(/\./g, '').replace(',', '.')) || 0;
                                const p2Num = parseFloat(p2Str.replace(/\./g, '').replace(',', '.')) || 0;
                                monthLineTotal += p1Num + p2Num;

                                combinedRows.push({
                                    CDOPERADOR: baseOp,
                                    DATA1: r1.DTOPERACAO || '',
                                    PASS1: r1.PASSAGEIROS || '',
                                    DATA2: r2.DTOPERACAO || '',
                                    PASS2: r2.PASSAGEIROS || ''
                                });
                            }

                            const passHeader = infoExtra.includes('EQUIVALENTE') ? 'PASS. EQUIV.' : 'PASSAGEIROS';

                            const columns = [
                                { header: 'OPERADOR', dataKey: 'CDOPERADOR' },
                                { header: '1ª QUINZENA', dataKey: 'DATA1' },
                                { header: passHeader, dataKey: 'PASS1' },
                                { header: '2ª QUINZENA', dataKey: 'DATA2' },
                                { header: passHeader, dataKey: 'PASS2' }
                            ];

                            const headerLeft = `LINHA ${lineGroup.linha} – ${lineGroup.nome}`;
                            const headerRight = monthNameFull.toUpperCase();
                            const headerTotal = `TOTAL: ${monthLineTotal.toLocaleString('pt-BR')}`;

                            doc.setFontSize(12);
                            doc.setFont("helvetica", "bold");

                            const pageWidth = doc.internal.pageSize.width;
                            const marginX = 18;

                            const dateWidth = doc.getTextWidth(headerRight);
                            const dateX = pageWidth - marginX;

                            const totalWidth = doc.getTextWidth(headerTotal);
                            const totalX = dateX - dateWidth - 15;

                            const lineX = marginX;

                            const totalStartX = totalX - totalWidth;
                            const maxNameWidth = totalStartX - lineX - 10;

                            let nameToPrint = headerLeft;
                            if (doc.getTextWidth(nameToPrint) > maxNameWidth) {
                                while (doc.getTextWidth(nameToPrint + "...") > maxNameWidth && nameToPrint.length > 0) {
                                    nameToPrint = nameToPrint.slice(0, -1);
                                }
                                nameToPrint += "...";
                            }

                            const pageHeight = doc.internal.pageSize.height;
                            const rowsCount = combinedRows.length;
                            const tableHeightApprox = 20 + (rowsCount * 7);

                            if (firstPageRendered && (startY + tableHeightApprox > pageHeight - marginBottom)) {
                                doc.addPage();
                                startY = headerHeight;
                            } else if (!firstPageRendered) {
                                startY = headerHeight;
                            }
                            firstPageRendered = true;

                            doc.setFillColor(lightGray[0], lightGray[1], lightGray[2]);
                            doc.roundedRect(14, startY, doc.internal.pageSize.width - 28, 10, 1, 1, 'F');

                            doc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);

                            doc.text(nameToPrint, lineX, startY + 7);
                            doc.text(headerTotal, totalX, startY + 7, { align: 'right' });
                            doc.text(headerRight, dateX, startY + 7, { align: 'right' });

                            startY += 11;

                            doc.autoTable({
                                columns: columns,
                                body: combinedRows,
                                startY: startY,
                                theme: 'grid',
                                styles: { fontSize: 9, textColor: darkText, cellPadding: 1.5, valign: 'middle', lineColor: [220, 220, 220], lineWidth: 0.1 },
                                headStyles: { fillColor: secondaryColor, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 10, halign: 'center', valign: 'middle' },
                                columnStyles: {
                                    0: { halign: 'center', fontStyle: 'bold', textColor: secondaryColor },
                                    1: { halign: 'center' },
                                    2: { halign: 'center', fontStyle: 'bold' },
                                    3: { halign: 'center' },
                                    4: { halign: 'center', fontStyle: 'bold' }
                                },
                                alternateRowStyles: { fillColor: [248, 249, 250] },
                                margin: { left: 14, right: 14, top: headerHeight, bottom: marginBottom },
                                didDrawPage: function (data) { drawHeader(data); drawFooter(data); },
                                didParseCell: function (data) {
                                    if (data.section === 'body' && data.column.index === 0) {
                                        if (data.cell.raw && data.cell.raw.toString().toUpperCase().includes('TOTAL')) {
                                            data.row.styles.fontStyle = 'bold';
                                            data.row.styles.fillColor = [255, 240, 230];
                                            data.row.styles.textColor = primaryColor;
                                        }
                                    }
                                }
                            });

                            startY = doc.lastAutoTable.finalY + 12;
                        }
                    }
                    else if (type === 'mean') {
                        const dataByMonth = {};
                        formattedBody.forEach(row => {
                            if (!row.DTOPERACAO) return;
                            const [day, month, year] = row.DTOPERACAO.split('/');
                            const key = `${year}${month}`;
                            if (!dataByMonth[key]) {
                                dataByMonth[key] = {
                                    monthIndex: parseInt(month, 10),
                                    year: parseInt(year, 10),
                                    rows: []
                                };
                            }
                            dataByMonth[key].rows.push(row);
                        });

                        const sortedMonthKeys = Object.keys(dataByMonth).sort((a, b) => parseInt(a) - parseInt(b));

                        for (let i = 0; i < sortedMonthKeys.length; i++) {
                            OuvidoriasApp.pageLogic.utils.setLoading(true, `Gerando PDF... Mês ${i + 1}/${sortedMonthKeys.length}`);
                            await new Promise(r => setTimeout(r, 0));

                            const monthKey = sortedMonthKeys[i];
                            const monthData = dataByMonth[monthKey];
                            const monthNameFull = `${getMonthName(monthData.monthIndex)} ${monthData.year}`;

                            if (firstPageRendered) {
                                doc.addPage();
                            }

                            startY = headerHeight + 15;

                            let columns = [
                                { header: 'OPERADOR', dataKey: 'CDOPERADOR' },
                                { header: 'CÓD. LINHA', dataKey: 'CDLINHA' },
                                { header: 'NOME DA LINHA', dataKey: 'LINHA_NOME' },
                                { header: 'DIAS', dataKey: 'DIAS_CONTABILIZADOS' },
                                { header: 'MÉDIA PASS.', dataKey: 'MÉDIA_PASSAGEIROS' }
                            ];

                            const hasEquivalentData = monthData.rows.some(d => d.MÉDIA_EQUIVALENTE !== undefined && d.MÉDIA_EQUIVALENTE !== '-');

                            if (infoExtra.includes('EQUIVALENTE') && hasEquivalentData) {
                                columns = columns.filter(c => c.dataKey !== 'MÉDIA_PASSAGEIROS');
                                columns.push({ header: 'MÉDIA EQUIV.', dataKey: 'MÉDIA_EQUIVALENTE' });
                            }

                            monthData.rows.sort((a, b) => {
                                if (a.CDOPERADOR !== b.CDOPERADOR) return a.CDOPERADOR.localeCompare(b.CDOPERADOR);
                                return parseInt(a.CDLINHA) - parseInt(b.CDLINHA);
                            });

                            doc.autoTable({
                                columns: columns,
                                body: monthData.rows,
                                startY: startY,
                                theme: 'striped',
                                styles: { fontSize: 10, textColor: darkText, cellPadding: 2, valign: 'middle', lineColor: [220, 220, 220], lineWidth: 0.1 },
                                headStyles: { fillColor: secondaryColor, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 11, halign: 'center', valign: 'middle' },
                                columnStyles: {
                                    0: { halign: 'center', fontStyle: 'bold', textColor: secondaryColor, width: 25 },
                                    1: { halign: 'center', fontStyle: 'bold', width: 25 },
                                    2: { halign: 'left' },
                                    3: { halign: 'center', width: 20 },
                                    4: { halign: 'center', fontStyle: 'bold', width: 30 },
                                    5: { halign: 'center', width: 30 }
                                },
                                margin: { left: 14, right: 14, top: headerHeight + 15, bottom: marginBottom },
                                didDrawPage: function (data) {
                                    drawHeader(data);
                                    drawFooter(data);
                                    const currentDoc = data.doc;
                                    currentDoc.setFontSize(12);
                                    currentDoc.setFont("helvetica", "bold");
                                    currentDoc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
                                    currentDoc.text(`MÊS DE REFERÊNCIA: ${monthNameFull}`, 14, headerHeight + 10);
                                }
                            });

                            startY = doc.lastAutoTable.finalY + 15;
                            firstPageRendered = true;
                        }
                    }

                    const finalName = fileNameBase.toLowerCase().endsWith('.pdf') ? fileNameBase : `${fileNameBase.replace('.xlsx', '')}.pdf`;
                    doc.save(finalName);
                    return true;
                },

                async findDateRangeInFile(file) {
                    return new Promise((resolve, reject) => {
                        if (!file) return resolve(null);
                        const reader = new FileReader();
                        reader.onload = (event) => {
                            try {
                                const text = event.target.result;
                                const lines = text.split(/\r\n|\n/).filter(Boolean);
                                if (lines.length < 2) return resolve(null);
                                const delimiters = [';', '\t', ','];
                                let delimiter = ';';
                                let maxCols = 0;
                                for (const d of delimiters) {
                                    const cols = lines[0].split(d).length;
                                    if (cols > maxCols) { maxCols = cols; delimiter = d; }
                                }
                                const header = lines[0].split(delimiter).map(h => h.trim().toUpperCase());
                                const dateColIndex = header.indexOf('DTOPERACAO');
                                if (dateColIndex === -1) return resolve(null);
                                let minDateNum = Infinity;
                                let maxDateNum = 0;
                                for (let i = 1; i < lines.length; i++) {
                                    const parts = lines[i].split(delimiter);
                                    if (parts.length <= dateColIndex) continue;
                                    const dateStr = parts[dateColIndex]?.trim();
                                    const match = dateStr?.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
                                    if (match) {
                                        const [_, day, month, year] = match;
                                        const dateNum = parseInt(`${year}${month}${day}`);
                                        if (dateNum < minDateNum) minDateNum = dateNum;
                                        if (dateNum > maxDateNum) maxDateNum = dateNum;
                                    }
                                }
                                if (minDateNum === Infinity || maxDateNum === 0) return resolve(null);
                                const minStr = minDateNum.toString();
                                const maxStr = maxDateNum.toString();
                                const minDate = new Date(Date.UTC(minStr.substring(0, 4), minStr.substring(4, 6) - 1, minStr.substring(6, 8)));
                                const maxDate = new Date(Date.UTC(maxStr.substring(0, 4), maxStr.substring(4, 6) - 1, maxStr.substring(6, 8)));
                                resolve({ minDate, maxDate });
                            } catch (e) { reject(e); }
                        };
                        reader.onerror = () => reject(new Error(`Erro ao ler arquivo ${file.name}`));
                        reader.readAsText(file, 'latin1');
                    });
                },
            },

            ui: {
                switchAssistant(assistantToShow) {
                    const { demandLineContainer, demandMeanContainer, selectDemandLineBtn, selectDemandMeanBtn } = OuvidoriasApp.elements;
                    if (assistantToShow === 'line') {
                        demandLineContainer.classList.remove('hidden');
                        demandMeanContainer.classList.add('hidden');
                        selectDemandLineBtn.classList.add('active');
                        selectDemandMeanBtn.classList.remove('active');
                    } else if (assistantToShow === 'mean') {
                        demandLineContainer.classList.add('hidden');
                        demandMeanContainer.classList.remove('hidden');
                        selectDemandLineBtn.classList.remove('active');
                        selectDemandMeanBtn.classList.add('active');
                    }
                },

                setupAppUI(processType) {
                    OuvidoriasApp.state.lineProcessType = processType;
                    OuvidoriasApp.state.lineConfig = OuvidoriasApp.config.lineConfigs[processType];
                    const { assistantTitle, validEmpresasInfo, fileInputPermContainer, fileInputConcContainer, dropZonePerm, lineTipoDia, lineProcessEquivalentContainer } = OuvidoriasApp.elements;

                    assistantTitle.textContent = OuvidoriasApp.state.lineConfig.title;
                    validEmpresasInfo.textContent = `Válidas: ${OuvidoriasApp.state.lineConfig.empresas.join(', ')}`;

                    const dropZonePermText = dropZonePerm.querySelector('.drop-zone-text');

                    if (processType === 'stpp_rmr') {
                        fileInputPermContainer.classList.remove('hidden');
                        fileInputConcContainer.classList.remove('hidden');
                        dropZonePermText.innerHTML = `Arraste o arquivo TXT das <strong>Permissionárias</strong> aqui ou clique para selecionar`;
                    } else if (processType === 'concessionarias') {
                        fileInputPermContainer.classList.remove('hidden');
                        fileInputConcContainer.classList.remove('hidden');
                        dropZonePermText.innerHTML = `Arraste o arquivo TXT das <strong>Permissionárias</strong> (Ref. Dia Útil) ou clique`;
                    } else {
                        fileInputPermContainer.classList.remove('hidden');
                        fileInputConcContainer.classList.add('hidden');
                        dropZonePermText.innerHTML = `Arraste o arquivo TXT das <strong>Permissionárias</strong> aqui ou clique para selecionar`;
                    }

                    lineTipoDia.innerHTML = '';
                    OuvidoriasApp.config.tiposDiaComuns.forEach(opt => {
                        const option = document.createElement('option');
                        option.value = opt;
                        option.textContent = opt;
                        lineTipoDia.appendChild(option);
                    });
                    lineTipoDia.value = "TODOS";

                    if (lineProcessEquivalentContainer) {
                        lineProcessEquivalentContainer.classList.remove('hidden');
                    }

                    OuvidoriasApp.elements.selectionScreen.classList.add('hidden');
                    OuvidoriasApp.elements.mainAssistant.classList.remove('hidden');
                },

                setupMeanAppUI(processType) {
                    OuvidoriasApp.state.meanProcessType = processType;
                    OuvidoriasApp.state.meanConfig = OuvidoriasApp.config.meanConfigs[processType];
                    OuvidoriasApp.pageLogic.ui.clearMeanFields(false);
                    OuvidoriasApp.elements.meanAssistantTitle.textContent = OuvidoriasApp.state.meanConfig.title;
                    OuvidoriasApp.elements.meanValidEmpresasInfo.textContent = `Válidas: ${OuvidoriasApp.state.meanConfig.empresas.join(', ')}`;
                    const fileContainer = OuvidoriasApp.elements.meanFileInputsContainer;
                    fileContainer.innerHTML = '';
                    OuvidoriasApp.state.meanConfig.fileConfigs.forEach(fileConf => {
                        const inputId = `mean-file-input-${fileConf.id}`;
                        const wrapper = document.createElement('div');
                        wrapper.className = 'file-input-wrapper';
                        wrapper.innerHTML = `
                            <label class="form-grid-label text-sm">${fileConf.label}:</label>
                            <div class="flex-grow flex items-center gap-2">
                                <span id="file-info-${fileConf.id}" class="file-info-text">Nenhum arquivo selecionado</span>
                                <button id="browse-btn-${fileConf.id}" type="button" class="btn btn-secondary !py-2 !px-3 text-sm">Procurar...</button>
                                <input type="file" id="${inputId}" class="hidden" accept=".txt">
                            </div>
                        `;
                        fileContainer.appendChild(wrapper);
                        document.getElementById(`browse-btn-${fileConf.id}`).addEventListener('click', () => document.getElementById(inputId).click());
                        document.getElementById(inputId).addEventListener('change', (e) => {
                            if (e.target.files.length > 0) {
                                const file = e.target.files[0];
                                if (file.name.toLowerCase().endsWith('.txt')) {
                                    OuvidoriasApp.state.meanFiles[fileConf.id] = file;
                                    const infoText = document.getElementById(`file-info-${fileConf.id}`);
                                    infoText.textContent = file.name;
                                    infoText.classList.add('loaded');
                                } else {
                                    OuvidoriasApp.pageLogic.utils.showFlashMessage('Por favor, selecione um arquivo no formato .txt', 'error');
                                    e.target.value = '';
                                }
                            }
                        });
                    });
                    const tipoDiaSelect = OuvidoriasApp.elements.meanTipoDia;
                    tipoDiaSelect.innerHTML = '';
                    OuvidoriasApp.state.meanConfig.tiposDia.forEach(opt => {
                        const option = document.createElement('option');
                        option.value = opt;
                        option.textContent = opt;
                        tipoDiaSelect.appendChild(option);
                    });
                    tipoDiaSelect.value = OuvidoriasApp.state.meanConfig.tiposDia[0];
                    const equivalentContainer = OuvidoriasApp.elements.meanProcessEquivalentContainer;

                    equivalentContainer.classList.remove('hidden');

                    OuvidoriasApp.elements.meanSelectionScreen.classList.add('hidden');
                    OuvidoriasApp.elements.mainMeanAssistant.classList.remove('hidden');
                },

                setupDropZone(dropZoneId, fileInputId, fileStateKey) {
                    const dropZone = document.getElementById(dropZoneId);
                    if (!dropZone) return;
                    const fileInput = document.getElementById(fileInputId);
                    const dropZoneText = dropZone.querySelector('.drop-zone-text');
                    const originalTextHTML = dropZoneText.innerHTML;

                    const handleFile = (file) => {
                        if (file && file.name.toLowerCase().endsWith('.txt')) {
                            OuvidoriasApp.state[fileStateKey] = file;
                            dropZone.classList.add('file-loaded');
                            dropZoneText.textContent = file.name;
                            OuvidoriasApp.pageLogic.handlers.handleAutoDetectPeriod();
                        } else {
                            OuvidoriasApp.pageLogic.utils.showFlashMessage('Por favor, selecione um arquivo no formato .txt', 'error');
                        }
                    };
                    dropZone.addEventListener('click', () => fileInput.click());
                    fileInput.addEventListener('change', () => { if (fileInput.files.length > 0) handleFile(fileInput.files[0]); });
                    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
                    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
                    dropZone.addEventListener('drop', (e) => {
                        e.preventDefault();
                        dropZone.classList.remove('dragover');
                        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
                    });
                    dropZone.reset = () => {
                        OuvidoriasApp.state[fileStateKey] = null;
                        fileInput.value = '';
                        dropZone.classList.remove('file-loaded');
                        dropZoneText.innerHTML = originalTextHTML;
                    };
                },

                renderPeriods() {
                    const { periodListContainer } = OuvidoriasApp.elements;
                    periodListContainer.innerHTML = '';
                    OuvidoriasApp.state.linePeriods.forEach((period, index) => {
                        const item = document.createElement('div');
                        item.className = 'period-item';
                        if (index === OuvidoriasApp.state.lineSelectedPeriodIndex) {
                            item.classList.add('selected');
                        }
                        item.textContent = period.display;
                        item.dataset.index = index;
                        item.addEventListener('click', () => {
                            OuvidoriasApp.state.lineSelectedPeriodIndex = (OuvidoriasApp.state.lineSelectedPeriodIndex === index) ? -1 : index;
                            OuvidoriasApp.pageLogic.ui.renderPeriods();
                        });
                        periodListContainer.appendChild(item);
                    });
                },

                clearAllFields() {
                    OuvidoriasApp.elements.dropZonePerm?.reset();
                    OuvidoriasApp.elements.dropZoneConc?.reset();
                    OuvidoriasApp.elements.codLinhas.value = '';
                    OuvidoriasApp.elements.empresas.value = '';
                    OuvidoriasApp.elements.ouvidoriaId.value = 'Resp_Ouvidoria_';
                    OuvidoriasApp.elements.dataInicio.value = '';
                    OuvidoriasApp.elements.dataFim.value = '';
                    OuvidoriasApp.elements.exportFormatLine.checked = false;
                    OuvidoriasApp.state.linePeriods = [];
                    OuvidoriasApp.state.lineSelectedPeriodIndex = -1;
                    OuvidoriasApp.elements.lineTipoDia.value = "TODOS";
                    OuvidoriasApp.elements.lineProcessEquivalent.checked = false;
                    OuvidoriasApp.pageLogic.ui.renderPeriods();
                },

                clearMeanFields(fullReset = true) {
                    if (fullReset) {
                        OuvidoriasApp.state.meanProcessType = null;
                        OuvidoriasApp.state.meanConfig = null;
                        OuvidoriasApp.elements.meanFileInputsContainer.innerHTML = '';
                    } else {
                        if (OuvidoriasApp.state.meanConfig && OuvidoriasApp.state.meanConfig.fileConfigs) {
                            OuvidoriasApp.state.meanConfig.fileConfigs.forEach(conf => {
                                const input = document.getElementById(`mean-file-input-${conf.id}`);
                                const info = document.getElementById(`file-info-${conf.id}`);
                                if (input) input.value = '';
                                if (info) {
                                    info.textContent = 'Nenhum arquivo selecionado';
                                    info.classList.remove('loaded');
                                }
                            });
                        }
                    }
                    OuvidoriasApp.state.meanFiles = {};
                    OuvidoriasApp.elements.meanCodLinhas.value = '';
                    OuvidoriasApp.elements.meanEmpresas.value = '';
                    OuvidoriasApp.elements.meanDataInicio.value = '';
                    OuvidoriasApp.elements.meanDataFim.value = '';
                    OuvidoriasApp.elements.meanSpecificMonths.value = '';
                    OuvidoriasApp.elements.meanMultiMonthToggle.checked = false;
                    OuvidoriasApp.elements.meanDateRangeContainer.classList.remove('hidden');
                    OuvidoriasApp.elements.meanSpecificMonthsContainer.classList.add('hidden');
                    OuvidoriasApp.elements.meanProcessEquivalent.checked = false;
                    OuvidoriasApp.elements.exportFormatMean.checked = false;
                    OuvidoriasApp.elements.meanOuvidoriaId.value = 'Resp_Ouvidoria_';
                },
            },

            handlers: {
                initAssistantSwitcher() {
                    OuvidoriasApp.elements.selectDemandLineBtn.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.switchAssistant('line'));
                    OuvidoriasApp.elements.selectDemandMeanBtn.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.switchAssistant('mean'));
                },

                initLineDemandHandlers() {
                    const { selectPermissionarias, selectConcessionarias, selectStppRmr, backToSelectionBtn, addPeriodBtn, removePeriodBtn, clearBtn, processBtn } = OuvidoriasApp.elements;
                    selectPermissionarias.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.setupAppUI('permissionarias'));
                    selectConcessionarias.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.setupAppUI('concessionarias'));
                    selectStppRmr.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.setupAppUI('stpp_rmr'));
                    backToSelectionBtn.addEventListener('click', () => {
                        OuvidoriasApp.elements.mainAssistant.classList.add('hidden');
                        OuvidoriasApp.elements.selectionScreen.classList.remove('hidden');
                        OuvidoriasApp.pageLogic.ui.clearAllFields();
                    });
                    OuvidoriasApp.pageLogic.ui.setupDropZone('drop-zone-perm', 'file-input-perm', 'linePermFile');
                    OuvidoriasApp.pageLogic.ui.setupDropZone('drop-zone-conc', 'file-input-conc', 'lineConcFile');
                    addPeriodBtn.addEventListener('click', OuvidoriasApp.pageLogic.handlers.handleAddPeriod);
                    removePeriodBtn.addEventListener('click', OuvidoriasApp.pageLogic.handlers.handleRemovePeriod);
                    clearBtn.addEventListener('click', OuvidoriasApp.pageLogic.ui.clearAllFields);
                    processBtn.addEventListener('click', OuvidoriasApp.pageLogic.handlers.handleProcessDemandLine);
                },

                initMeanDemandHandlers() {
                    const { selectMeanPermissionarias, selectMeanConcessionarias, selectMeanStppRmr, backToMeanSelectionBtn, meanMultiMonthToggle, meanClearBtn, meanProcessBtn } = OuvidoriasApp.elements;
                    selectMeanPermissionarias.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.setupMeanAppUI('permissionarias'));
                    selectMeanConcessionarias.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.setupMeanAppUI('concessionarias'));
                    selectMeanStppRmr.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.setupMeanAppUI('stpp_rmr'));
                    backToMeanSelectionBtn.addEventListener('click', () => {
                        OuvidoriasApp.elements.mainMeanAssistant.classList.add('hidden');
                        OuvidoriasApp.elements.meanSelectionScreen.classList.remove('hidden');
                        OuvidoriasApp.pageLogic.ui.clearMeanFields(true);
                    });
                    meanMultiMonthToggle.addEventListener('change', (e) => {
                        const isChecked = e.target.checked;
                        OuvidoriasApp.elements.meanDateRangeContainer.classList.toggle('hidden', isChecked);
                        OuvidoriasApp.elements.meanSpecificMonthsContainer.classList.toggle('hidden', !isChecked);
                    });
                    meanClearBtn.addEventListener('click', () => OuvidoriasApp.pageLogic.ui.clearMeanFields(false));
                    meanProcessBtn.addEventListener('click', OuvidoriasApp.pageLogic.handlers.handleProcessDemandMean);
                },

                async handleAutoDetectPeriod() {
                    const filesToScan = [OuvidoriasApp.state.linePermFile, OuvidoriasApp.state.lineConcFile].filter(Boolean);
                    if (filesToScan.length === 0) return;
                    OuvidoriasApp.pageLogic.utils.setLoading(true, 'Analisando datas...');
                    try {
                        const ranges = await Promise.all(filesToScan.map(file => OuvidoriasApp.pageLogic.utils.findDateRangeInFile(file)));
                        const validRanges = ranges.filter(Boolean);
                        if (validRanges.length === 0) {
                            OuvidoriasApp.pageLogic.utils.showFlashMessage('Não foi possível detectar um intervalo de datas nos arquivos. Por favor, insira o período manualmente.', 'warning');
                            return;
                        }
                        let overallMinDate = validRanges[0].minDate;
                        let overallMaxDate = validRanges[0].maxDate;
                        for (let i = 1; i < validRanges.length; i++) {
                            if (validRanges[i].minDate < overallMinDate) overallMinDate = validRanges[i].minDate;
                            if (validRanges[i].maxDate > overallMaxDate) overallMaxDate = validRanges[i].maxDate;
                        }
                        const inicioStr = overallMinDate.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
                        const fimStr = overallMaxDate.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
                        const display = `${inicioStr} - ${fimStr}`;
                        const inicioNum = parseInt(`${overallMinDate.getUTCFullYear()}${(overallMinDate.getUTCMonth() + 1).toString().padStart(2, '0')}${overallMinDate.getUTCDate().toString().padStart(2, '0')}`);
                        const fimNum = parseInt(`${overallMaxDate.getUTCFullYear()}${(overallMaxDate.getUTCMonth() + 1).toString().padStart(2, '0')}${overallMaxDate.getUTCDate().toString().padStart(2, '0')}`);
                        OuvidoriasApp.state.linePeriods = [{
                            inicio: overallMinDate,
                            fim: overallMaxDate,
                            display,
                            inicioNum,
                            fimNum
                        }];
                        OuvidoriasApp.pageLogic.ui.renderPeriods();
                    } catch (error) {
                        console.error('Erro ao detectar período:', error);
                        OuvidoriasApp.pageLogic.utils.showFlashMessage(`Ocorreu um erro ao analisar as datas do arquivo: ${error.message}`, 'error');
                    } finally {
                        OuvidoriasApp.pageLogic.utils.setLoading(false);
                    }
                },

                handleAddPeriod() {
                    const { dataInicio, dataFim } = OuvidoriasApp.elements;
                    const inicioStr = dataInicio.value;
                    const fimStr = dataFim.value;
                    const dateRegex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
                    if (!dateRegex.test(inicioStr) || !dateRegex.test(fimStr)) {
                        return OuvidoriasApp.pageLogic.utils.showFlashMessage("Formato de data inválido. Use DD/MM/AAAA.", 'error');
                    }
                    const [, d1, m1, y1] = inicioStr.match(dateRegex);
                    const [, d2, m2, y2] = fimStr.match(dateRegex);
                    const inicioObj = new Date(Date.UTC(y1, m1 - 1, d1));
                    const fimObj = new Date(Date.UTC(y2, m2 - 1, d2));
                    if (fimObj < inicioObj) {
                        return OuvidoriasApp.pageLogic.utils.showFlashMessage("A 'Data Fim' deve ser maior ou igual à 'Data Início'.", 'error');
                    }
                    const inicioNum = parseInt(`${y1}${(m1).padStart(2, '0')}${d1.padStart(2, '0')}`);
                    const fimNum = parseInt(`${y2}${(m2).padStart(2, '0')}${d2.padStart(2, '0')}`);
                    const display = `${inicioStr} - ${fimStr}`;
                    if (OuvidoriasApp.state.linePeriods.some(p => p.display === display)) {
                        return OuvidoriasApp.pageLogic.utils.showFlashMessage("Este período já foi adicionado.", 'info');
                    }
                    OuvidoriasApp.state.linePeriods.push({ inicio: inicioObj, fim: fimObj, display, inicioNum, fimNum });
                    dataInicio.value = '';
                    dataFim.value = '';
                    OuvidoriasApp.pageLogic.ui.renderPeriods();
                },

                handleRemovePeriod() {
                    if (OuvidoriasApp.state.lineSelectedPeriodIndex > -1) {
                        OuvidoriasApp.state.linePeriods.splice(OuvidoriasApp.state.lineSelectedPeriodIndex, 1);
                        OuvidoriasApp.state.lineSelectedPeriodIndex = -1;
                        OuvidoriasApp.pageLogic.ui.renderPeriods();
                    } else {
                        OuvidoriasApp.pageLogic.utils.showFlashMessage("Nenhum período selecionado para remover.", 'info');
                    }
                },

                async handleProcessDemandLine() {
                    const { validation, data } = OuvidoriasApp.pageLogic.processing.validateInputs();
                    if (!validation.isValid) {
                        return OuvidoriasApp.pageLogic.utils.showFlashMessage(validation.message, 'error');
                    }
                    const usePDF = OuvidoriasApp.elements.exportFormatLine.checked;
                    OuvidoriasApp.pageLogic.utils.setLoading(true, usePDF ? 'Gerando PDF...' : 'Gerando Excel...');
                    try {
                        const finalData = await OuvidoriasApp.pageLogic.processing.processData(data);
                        const usePDF = OuvidoriasApp.elements.exportFormatLine.checked;

                        const extraInfo = {
                            tipoDia: data.tipoDia,
                            equivalente: data.useEquivalent
                        };

                        if (usePDF) {
                            let infoStr = `Período: ${data.periods[0].display}`;
                            if (extraInfo.tipoDia) infoStr += ` | ${extraInfo.tipoDia}`;
                            if (extraInfo.equivalente) infoStr += ` | PASS. EQUIVALENTE`;

                            // Load company logos
                            OuvidoriasApp.pageLogic.utils.setLoading(true, 'Carregando logos...');
                            const logoFiles = {
                                'BOA': '/static/images/borborema.png',
                                'CAX': '/static/images/caxanga.png',
                                'CNO': '/static/images/conorte.png',
                                'CSR': '/static/images/consorcio_recife.png',
                                'EME': '/static/images/empresa_metropolitana.png',
                                'GLO': '/static/images/globo.png',
                                'MOB': '/static/images/mobi.png',
                                'SJT': '/static/images/sao_judas_tadeu.png',
                                'VML': '/static/images/viacao_mirim.png'
                            };

                            const loadedLogos = {};
                            const promises = Object.entries(logoFiles).map(async ([key, url]) => {
                                const b64 = await OuvidoriasApp.pageLogic.utils.loadImageBase64(url, 200);
                                if (b64) loadedLogos[key] = b64;
                            });
                            await Promise.all(promises);

                            await OuvidoriasApp.pageLogic.utils.generatePDF(finalData, OuvidoriasApp.state.lineConfig.title, data.fileName, 'line', infoStr, loadedLogos);
                            OuvidoriasApp.pageLogic.utils.showFlashMessage(`Relatório PDF "${data.fileName}" gerado com sucesso.`, 'success');
                        } else {
                            const ctmLogoBase64 = await OuvidoriasApp.pageLogic.utils.loadImageBase64('/static/images/ctm.png');
                            await OuvidoriasApp.pageLogic.utils.createExcelWorkbook(finalData, data.periods, data.fileName, { ctm: ctmLogoBase64 }, extraInfo);
                            OuvidoriasApp.pageLogic.utils.showFlashMessage(`Relatório Excel "${data.fileName}" gerado.`, 'success');
                        }
                    } catch (error) {
                        console.error("Erro no processamento:", error);
                        OuvidoriasApp.pageLogic.utils.showFlashMessage(`Ocorreu um erro: ${error.message}`, 'error', 0);
                    } finally {
                        OuvidoriasApp.pageLogic.utils.setLoading(false);
                    }
                },

                async handleProcessDemandMean() {
                    const { isValid, message, data } = OuvidoriasApp.pageLogic.processing.validateMeanInputs();
                    if (!isValid) {
                        return OuvidoriasApp.pageLogic.utils.showFlashMessage(message, 'error');
                    }
                    const usePDF = OuvidoriasApp.elements.exportFormatMean.checked;
                    OuvidoriasApp.pageLogic.utils.setLoading(true, usePDF ? 'Gerando PDF...' : 'Gerando Excel...');
                    try {
                        let allData = [];
                        let diaTipoMap = new Map();
                        if (data.files.permLookupFile || (OuvidoriasApp.state.meanProcessType === 'permissionarias' || OuvidoriasApp.state.meanProcessType === 'stpp_rmr')) {
                            const fileForMap = data.files.permLookupFile || data.files.permFile;
                            const mapData = await OuvidoriasApp.pageLogic.processing.parseMeanTxtFile(fileForMap, ['DTOPERACAO', 'DSDIATIPO']);
                            mapData.forEach(row => {
                                const match = row.DTOPERACAO.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
                                if (match) {
                                    const dateKey = `${match[3]}-${match[2]}-${match[1]}`;
                                    if (!diaTipoMap.has(dateKey)) {
                                        diaTipoMap.set(dateKey, row.DSDIATIPO.toUpperCase());
                                    }
                                }
                            });
                        }

                        if (OuvidoriasApp.state.meanProcessType === 'permissionarias' || OuvidoriasApp.state.meanProcessType === 'stpp_rmr') {
                            const permData = await OuvidoriasApp.pageLogic.processing.parseMeanTxtFile(data.files.permFile, OuvidoriasApp.config.meanConfigs.permissionarias.fileConfigs[0].cols);
                            permData.forEach(row => {
                                row.NMPASSTOTAL = row.NMEFETPASST;
                                row.NMPASSEQUIVALENTE = row.NMEFETPASST;
                            });
                            allData.push(...permData);
                        }
                        if (OuvidoriasApp.state.meanProcessType === 'concessionarias' || OuvidoriasApp.state.meanProcessType === 'stpp_rmr') {
                            const concData = await OuvidoriasApp.pageLogic.processing.parseMeanTxtFile(data.files.concFile, OuvidoriasApp.config.meanConfigs.concessionarias.fileConfigs[0].cols);
                            allData.push(...concData);
                        }

                        const allRawData = allData.map(row => {
                            const match = row.DTOPERACAO.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
                            if (!match) return null;
                            const [_, day, month, year] = match;
                            const dateObj = new Date(Date.UTC(year, month - 1, day));
                            const dateKey = `${year}-${month}-${day}`;
                            row.date = dateObj;
                            if (!row.DSDIATIPO && diaTipoMap.has(dateKey)) {
                                row.DSDIATIPO = diaTipoMap.get(dateKey);
                            }
                            row.DSDIATIPO = (row.DSDIATIPO || '').toUpperCase();

                            row.passTotal = parseFloat((row.NMPASSTOTAL || '0').replace(/\./g, '').replace(',', '.')) || 0;
                            row.passEquiv = parseFloat((row.NMPASSEQUIVALENTE || '0').replace(/\./g, '').replace(',', '.')) || 0;

                            return row;
                        }).filter(Boolean);

                        const usePDF = OuvidoriasApp.elements.exportFormatMean.checked;
                        if (data.isMultiMonth) {
                            if (usePDF) {
                                let combinedExcelData = [];
                                for (const period of data.periods) {
                                    const singleMonthInputs = { ...data, periods: [period] };
                                    const excelData = await OuvidoriasApp.pageLogic.processing.processMeanData(singleMonthInputs, allRawData);
                                    const refDate = `01/${period.monthStr.replace('-', '/')}`;
                                    excelData.forEach(r => r.DTOPERACAO = refDate);
                                    combinedExcelData = [...combinedExcelData, ...excelData];
                                }
                                if (combinedExcelData.length > 0) {
                                    let infoExtra = `Média (${data.tipoDia})`;
                                    if (data.processEquivalent) infoExtra += ` - INCLUI EQUIVALENTE`;

                                    // Load company logos
                                    OuvidoriasApp.pageLogic.utils.setLoading(true, 'Carregando logos...');
                                    const logoFiles = {
                                        'BOA': '/static/images/borborema.png',
                                        'CAX': '/static/images/caxanga.png',
                                        'CNO': '/static/images/conorte.png',
                                        'CSR': '/static/images/consorcio_recife.png',
                                        'EME': '/static/images/empresa_metropolitana.png',
                                        'GLO': '/static/images/globo.png',
                                        'MOB': '/static/images/mobi.png',
                                        'SJT': '/static/images/sao_judas_tadeu.png',
                                        'VML': '/static/images/viacao_mirim.png'
                                    };
                                    const loadedLogos = {};
                                    const promises = Object.entries(logoFiles).map(async ([key, url]) => {
                                        const b64 = await OuvidoriasApp.pageLogic.utils.loadImageBase64(url, 200);
                                        if (b64) loadedLogos[key] = b64;
                                    });
                                    await Promise.all(promises);

                                    await OuvidoriasApp.pageLogic.utils.generatePDF(combinedExcelData, OuvidoriasApp.state.meanConfig.title, data.fileName, 'mean', infoExtra, loadedLogos);
                                    OuvidoriasApp.pageLogic.utils.showFlashMessage(`Relatório PDF "${data.fileName}" gerado com sucesso.`, 'success');
                                } else {
                                    OuvidoriasApp.pageLogic.utils.showFlashMessage('Nenhum dado encontrado para gerar o PDF.', 'warning');
                                }
                            } else {
                                const workbook = new ExcelJS.Workbook();
                                let hasAnyData = false;
                                const ctmLogoBase64 = await OuvidoriasApp.pageLogic.utils.loadImageBase64('/static/images/ctm.png');
                                for (const period of data.periods) {
                                    const singleMonthInputs = { ...data, periods: [period] };
                                    const excelData = await OuvidoriasApp.pageLogic.processing.processMeanData(singleMonthInputs, allRawData);
                                    if (excelData.length > 0) {
                                        hasAnyData = true;
                                        OuvidoriasApp.pageLogic.utils.addMeanWorksheetToWorkbook(workbook, period.monthStr, excelData, singleMonthInputs, period.monthStr.replace('-', '/'), { ctm: ctmLogoBase64 });
                                    }
                                }
                                if (hasAnyData) {
                                    await OuvidoriasApp.pageLogic.utils.saveWorkbookWithPicker(workbook, data.fileName);
                                    OuvidoriasApp.pageLogic.utils.showFlashMessage(`Relatório Excel "${data.fileName}" gerado com sucesso com abas para cada mês.`, 'success');
                                } else {
                                    OuvidoriasApp.pageLogic.utils.showFlashMessage('Nenhum dado encontrado com os filtros aplicados para os meses especificados. O relatório não será gerado.', 'warning');
                                }
                            }
                        } else {
                            const excelData = await OuvidoriasApp.pageLogic.processing.processMeanData(data, allRawData);
                            if (excelData.length === 0) {
                                OuvidoriasApp.pageLogic.utils.showFlashMessage('Nenhum dado encontrado com os filtros aplicados. O relatório não será gerado.', 'warning');
                            } else {
                                if (usePDF) {
                                    const refDate = data.periods[0].start.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
                                    excelData.forEach(r => r.DTOPERACAO = refDate);

                                    const periodStr = `${data.periods[0].start.toLocaleDateString('pt-BR', { timeZone: 'UTC' })} a ${data.periods[0].end.toLocaleDateString('pt-BR', { timeZone: 'UTC' })}`;
                                    let infoExtra = `Média (${data.tipoDia}) - ${periodStr}`;
                                    if (data.processEquivalent) infoExtra += ` - INCLUI EQUIVALENTE`;

                                    // Load company logos
                                    OuvidoriasApp.pageLogic.utils.setLoading(true, 'Carregando logos...');
                                    const logoFiles = {
                                        'BOA': '/static/images/borborema.png',
                                        'CAX': '/static/images/caxanga.png',
                                        'CNO': '/static/images/conorte.png',
                                        'CSR': '/static/images/consorcio_recife.png',
                                        'EME': '/static/images/empresa_metropolitana.png',
                                        'GLO': '/static/images/globo.png',
                                        'MOB': '/static/images/mobi.png',
                                        'SJT': '/static/images/sao_judas_tadeu.png',
                                        'VML': '/static/images/viacao_mirim.png'
                                    };
                                    const loadedLogos = {};
                                    // Use promise all
                                    const promises = Object.entries(logoFiles).map(async ([key, url]) => {
                                        const b64 = await OuvidoriasApp.pageLogic.utils.loadImageBase64(url, 200);
                                        if (b64) loadedLogos[key] = b64;
                                    });
                                    await Promise.all(promises);

                                    await OuvidoriasApp.pageLogic.utils.generatePDF(excelData, OuvidoriasApp.state.meanConfig.title, data.fileName, 'mean', infoExtra, loadedLogos);
                                    OuvidoriasApp.pageLogic.utils.showFlashMessage(`Relatório PDF "${data.fileName}" gerado com sucesso.`, 'success');
                                } else {
                                    const workbook = new ExcelJS.Workbook();
                                    const periodStr = `${data.periods[0].start.toLocaleDateString('pt-BR', { timeZone: 'UTC' })} a ${data.periods[0].end.toLocaleDateString('pt-BR', { timeZone: 'UTC' })}`;
                                    const ctmLogoBase64 = await OuvidoriasApp.pageLogic.utils.loadImageBase64('/static/images/ctm.png');
                                    OuvidoriasApp.pageLogic.utils.addMeanWorksheetToWorkbook(workbook, 'Média de Demanda', excelData, data, periodStr, { ctm: ctmLogoBase64 });
                                    await OuvidoriasApp.pageLogic.utils.saveWorkbookWithPicker(workbook, data.fileName);
                                    OuvidoriasApp.pageLogic.utils.showFlashMessage(`Relatório Excel "${data.fileName}" gerado com sucesso.`, 'success');
                                }
                            }
                        }
                    } catch (error) {
                        console.error("Erro no processamento da demanda média:", error);
                        OuvidoriasApp.pageLogic.utils.showFlashMessage(`Ocorreu um erro: ${error.message}`, 'error', 0);
                    } finally {
                        OuvidoriasApp.pageLogic.utils.setLoading(false);
                    }
                }
            },

            processing: {
                validateInputs() {
                    const { lineProcessType, lineConfig, linePermFile, lineConcFile, linePeriods } = OuvidoriasApp.state;
                    const { ouvidoriaId, codLinhas, empresas, lineTipoDia, lineProcessEquivalent } = OuvidoriasApp.elements;
                    const ouvidoriaIdVal = ouvidoriaId.value.trim();
                    const codLinhasVal = codLinhas.value.trim();
                    const empresasVal = empresas.value.trim().toUpperCase();

                    if (lineProcessType === 'stpp_rmr' && (!linePermFile || !lineConcFile)) return { validation: { isValid: false, message: "Para STPP/RMR, ambos os arquivos TXT são obrigatórios." } };
                    if (lineProcessType === 'concessionarias' && !lineConcFile) return { validation: { isValid: false, message: "O arquivo TXT das Concessionárias é obrigatório." } };
                    if (lineProcessType === 'permissionarias' && !linePermFile) return { validation: { isValid: false, message: "O arquivo TXT das Permissionárias é obrigatório." } };
                    if (!codLinhasVal) return { validation: { isValid: false, message: "'Códigos das Linhas' é obrigatório." } };
                    if (!empresasVal) return { validation: { isValid: false, message: "'Empresa(s)' é obrigatório." } };
                    if (!linePeriods.length) return { validation: { isValid: false, message: "Adicione pelo menos um período de análise." } };
                    if (!ouvidoriaIdVal) return { validation: { isValid: false, message: "'Nº Ouvidoria/ID' é obrigatório." } };

                    let empresasValidadas;
                    if (empresasVal === "TODAS") {
                        empresasValidadas = lineConfig.empresas;
                    } else {
                        const empresasInput = empresasVal.split(',').map(e => e.trim());
                        const invalidas = empresasInput.filter(e => !lineConfig.empresas.includes(e));
                        if (invalidas.length) return { validation: { isValid: false, message: `Empresa(s) inválida(s): ${invalidas.join(', ')}` } };
                        empresasValidadas = empresasInput;
                    }
                    return {
                        validation: { isValid: true },
                        data: {
                            permFile: linePermFile,
                            concFile: lineConcFile,
                            codLinhas: codLinhasVal,
                            empresas: empresasValidadas,
                            periods: linePeriods,
                            fileName: ouvidoriaIdVal.endsWith('.xlsx') ? ouvidoriaIdVal : `${ouvidoriaIdVal}.xlsx`,
                            tipoDia: lineTipoDia.value,
                            useEquivalent: lineProcessEquivalent.checked
                        }
                    };
                },

                async processData(data) {
                    let finalData = [];
                    const { lineConfig } = OuvidoriasApp.state;

                    let diaTipoMap = new Map();
                    if (data.permFile) {
                        const mapData = await this.parseTxtFile(data.permFile, ['DTOPERACAO', 'DSDIATIPO'], null, {});
                        mapData.forEach(row => {
                            const match = row.DTOPERACAO.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
                            if (match && row.DSDIATIPO) {
                                const dateKey = `${match[3]}-${match[2]}-${match[1]}`;
                                if (!diaTipoMap.has(dateKey)) {
                                    diaTipoMap.set(dateKey, row.DSDIATIPO.toUpperCase());
                                }
                            }
                        });
                    }

                    if (OuvidoriasApp.state.lineProcessType === 'stpp_rmr') {
                        const permCols = OuvidoriasApp.config.lineConfigs.permissionarias.colunasExcel;

                        const permPass = data.useEquivalent ?
                            OuvidoriasApp.config.lineConfigs.permissionarias.colunaEquivalente :
                            OuvidoriasApp.config.lineConfigs.permissionarias.colunaPassageiros;

                        const permData = await this.parseTxtFile(data.permFile, permCols, permPass, data, diaTipoMap);

                        const concCols = OuvidoriasApp.config.lineConfigs.concessionarias.colunasExcel;
                        const concPass = data.useEquivalent ? OuvidoriasApp.config.lineConfigs.concessionarias.colunaEquivalente : OuvidoriasApp.config.lineConfigs.concessionarias.colunaPassageiros;
                        const concData = await this.parseTxtFile(data.concFile, concCols, concPass, data, diaTipoMap);

                        concData.forEach(row => {
                            row.VALOR_EXIBIDO = row[concPass];
                        });
                        permData.forEach(row => {
                            row.VALOR_EXIBIDO = row[permPass];
                        });

                        finalData = [...permData, ...concData];
                    } else if (OuvidoriasApp.state.lineProcessType === 'concessionarias') {
                        let targetPassColumn = lineConfig.colunaPassageiros;
                        if (data.useEquivalent && lineConfig.colunaEquivalente) {
                            targetPassColumn = lineConfig.colunaEquivalente;
                        }
                        finalData = await this.parseTxtFile(data.concFile, lineConfig.colunasExcel, targetPassColumn, data, diaTipoMap);
                        finalData.forEach(row => {
                            row.VALOR_EXIBIDO = row[targetPassColumn];
                        });
                    } else {
                        let targetPassColumn = lineConfig.colunaPassageiros;
                        if (data.useEquivalent && lineConfig.colunaEquivalente) {
                            targetPassColumn = lineConfig.colunaEquivalente;
                        }
                        finalData = await this.parseTxtFile(data.permFile, lineConfig.colunasExcel, targetPassColumn, data, diaTipoMap);
                        finalData.forEach(row => {
                            row.VALOR_EXIBIDO = row[targetPassColumn];
                        });
                    }

                    if (finalData.length > 0) {
                        finalData.sort((a, b) => {
                            if (a.CDOPERADOR !== b.CDOPERADOR) return a.CDOPERADOR.localeCompare(b.CDOPERADOR);
                            const linhaA = parseInt(a.CDLINHA, 10) || 0;
                            const linhaB = parseInt(b.CDLINHA, 10) || 0;
                            if (linhaA !== linhaB) return linhaA - linhaB;
                            return a.DTOPERACAO_NUM - b.DTOPERACAO_NUM;
                        });
                    } else {
                        console.warn("Nenhum dado encontrado com os filtros aplicados.");
                    }
                    return finalData;
                },

                async loadServerLinesFile() {
                    const filePath = '/static/linhas.xlsx';
                    try {
                        const response = await fetch(filePath);
                        if (!response.ok) {
                            console.warn(`Arquivo não encontrado em ${filePath} (Erro ${response.status}). Verifique se moveu o arquivo para a pasta 'static'.`);
                            return;
                        }
                        const buffer = await response.arrayBuffer();
                        const workbook = new ExcelJS.Workbook();
                        await workbook.xlsx.load(buffer);
                        const worksheet = workbook.worksheets[0];
                        const linesMap = {};
                        let headerRow = null;
                        let colIndices = {};
                        worksheet.eachRow((row, rowNumber) => {
                            if (headerRow) return;
                            const values = Array.isArray(row.values) ? row.values.slice(1) : [];
                            const operatorIdx = values.findIndex(v => v && v.toString().toUpperCase().includes('OPERADOR'));
                            const codeIdx = values.findIndex(v => v && (v.toString().toUpperCase().includes('CÓDIGO LINHA') || v.toString().toUpperCase().includes('CODIGO LINHA')));
                            const nameIdx = values.findIndex(v => v && v.toString().toUpperCase().includes('NOME LINHA'));
                            if (operatorIdx !== -1 && codeIdx !== -1 && nameIdx !== -1) {
                                headerRow = rowNumber;
                                colIndices = {
                                    operador: operatorIdx + 1,
                                    codigo: codeIdx + 1,
                                    nome: nameIdx + 1
                                };
                            }
                        });
                        if (!headerRow) {
                            headerRow = 1;
                            colIndices = { operador: 1, codigo: 2, nome: 3 };
                        }
                        const knownShortCodes = ["BOA", "CAX", "CSR", "EME", "GLO", "SJT", "VML", "CNO", "MOB"];
                        worksheet.eachRow((row, rowNumber) => {
                            if (rowNumber <= headerRow) return;
                            const operadorRaw = row.getCell(colIndices.operador).value;
                            const codigoRaw = row.getCell(colIndices.codigo).value;
                            const nome = row.getCell(colIndices.nome).value;
                            if (operadorRaw && codigoRaw && nome) {
                                const opString = operadorRaw.toString().toUpperCase().trim();
                                let shortCode = knownShortCodes.find(sc => opString.includes(sc));
                                const finalOp = shortCode || opString;
                                const finalCode = parseInt(codigoRaw).toString();
                                const key = `${finalOp}_${finalCode}`;
                                linesMap[key] = nome.toString().trim();
                            }
                        });
                        OuvidoriasApp.state.lineLinesData = linesMap;
                        console.log(`Cadastro de Linhas carregado de ${filePath} com sucesso.`);
                    } catch (error) {
                        console.error("Erro ao carregar arquivo de linhas:", error);
                    }
                },

                parseTxtFile(file, colunasDesejadas, colunaPassageiros, filters, diaTipoMap = null) {
                    return new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onload = (event) => {
                            try {
                                const text = event.target.result;
                                const lines = text.split(/\r\n|\n/).filter(Boolean);
                                if (lines.length < 2) return resolve([]);
                                const delimiters = [';', '\t', ','];
                                let delimiter = ';';
                                let maxCols = 0;
                                for (const d of delimiters) {
                                    const cols = lines[0].split(d).length;
                                    if (cols > maxCols) { maxCols = cols; delimiter = d; }
                                }
                                const header = lines[0].split(delimiter).map(h => h.trim().toUpperCase());
                                const requiredColsUpper = colunasDesejadas.map(c => c.toUpperCase());
                                const colIndices = {};
                                for (const col of requiredColsUpper) {
                                    const index = header.indexOf(col);
                                    if (index === -1) {
                                        if (['NMPASSEQUIVALENTE', 'DSDIATIPO'].includes(col)) continue;
                                        throw new Error(`Coluna obrigatória não encontrada: ${col}.`);
                                    }
                                    colIndices[col] = index;
                                }
                                const data = lines.slice(1).map(line => {
                                    const parts = line.split(delimiter);
                                    if (parts.length < maxCols) return null;
                                    let row = {};
                                    for (const col in colIndices) { row[col] = parts[colIndices[col]]?.trim() || ''; }
                                    return row;
                                }).filter(Boolean)
                                    .map(row => {
                                        const dateStr = row.DTOPERACAO;
                                        const match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
                                        if (match) {
                                            const [_, day, month, year] = match;
                                            row.DTOPERACAO_NUM = parseInt(`${year}${month.padStart(2, '0')}${day.padStart(2, '0')}`);
                                            row.DTOPERACAO_OBJ = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));

                                            if ((!row.DSDIATIPO || row.DSDIATIPO === '') && diaTipoMap) {
                                                const dateKey = `${year}-${month}-${day}`;
                                                if (diaTipoMap.has(dateKey)) {
                                                    row.DSDIATIPO = diaTipoMap.get(dateKey);
                                                }
                                            }
                                        }
                                        return row;
                                    })
                                    .filter(row => {
                                        if (!filters || Object.keys(filters).length === 0) return true;

                                        if (!filters.periods || filters.periods.length === 0) return !!row.DTOPERACAO_NUM;
                                        if (!row.DTOPERACAO_NUM || !filters.periods.some(p => row.DTOPERACAO_NUM >= p.inicioNum && row.DTOPERACAO_NUM <= p.fimNum)) return false;

                                        if (filters.codLinhas && filters.codLinhas.toUpperCase() !== 'TODAS' && !filters.codLinhas.split(',').map(l => l.trim()).includes(row.CDLINHA)) return false;

                                        if (filters.empresas && !filters.empresas.includes(row.CDOPERADOR.toUpperCase())) return false;

                                        if (!filters.tipoDia || filters.tipoDia === 'TODOS') return true;
                                        return row.DSDIATIPO && row.DSDIATIPO.toUpperCase() === filters.tipoDia;
                                    })
                                    .map(row => {
                                        if (colunaPassageiros) {
                                            const passColUpper = colunaPassageiros.toUpperCase();
                                            if (row[passColUpper]) {
                                                const numStr = row[passColUpper].replace(/\./g, '').replace(',', '.');
                                                row[passColUpper] = parseFloat(numStr) || 0;
                                            }
                                        }
                                        return row;
                                    });
                                resolve(data);
                            } catch (e) { reject(e); }
                        };
                        reader.onerror = () => reject(new Error(`Erro ao ler o arquivo ${file.name}`));
                        reader.readAsText(file, 'latin1');
                    });
                },

                validateMeanInputs() {
                    const { meanProcessType, meanConfig, meanFiles } = OuvidoriasApp.state;
                    for (const fileConf of meanConfig.fileConfigs) {
                        if (fileConf.required && !meanFiles[fileConf.id]) {
                            return { isValid: false, message: `O arquivo "${fileConf.label}" é obrigatório.` };
                        }
                    }
                    const codLinhas = OuvidoriasApp.elements.meanCodLinhas.value.trim();
                    if (!codLinhas) return { isValid: false, message: "'Códigos das Linhas' é obrigatório." };
                    const empresas = OuvidoriasApp.elements.meanEmpresas.value.trim().toUpperCase();
                    if (!empresas) return { isValid: false, message: "'Empresa(s)' é obrigatório." };
                    let empresasValidadas;
                    if (empresas === "TODAS") {
                        empresasValidadas = meanConfig.empresas;
                    } else {
                        const empresasInput = empresas.split(',').map(e => e.trim());
                        const invalidas = empresasInput.filter(e => !meanConfig.empresas.includes(e));
                        if (invalidas.length) return { isValid: false, message: `Empresa(s) inválida(s) para este tipo de relatório: ${invalidas.join(', ')}` };
                        empresasValidadas = empresasInput;
                    }
                    const isMultiMonth = OuvidoriasApp.elements.meanMultiMonthToggle.checked;
                    let periods = [];
                    let specificMonthsStr = '';
                    if (isMultiMonth) {
                        const monthsStr = OuvidoriasApp.elements.meanSpecificMonths.value.trim();
                        if (!monthsStr) return { isValid: false, message: "Pelo menos um mês deve ser informado no formato MM/AAAA." };
                        specificMonthsStr = monthsStr;
                        const monthsArr = monthsStr.split(',').map(m => m.trim());
                        for (const monthStr of monthsArr) {
                            const match = monthStr.match(/^(\d{2})\/(\d{4})$/);
                            if (!match) return { isValid: false, message: `Formato de mês inválido: "${monthStr}". Use MM/AAAA.` };
                            const month = parseInt(match[1], 10);
                            const year = parseInt(match[2], 10);
                            if (month < 1 || month > 12) return { isValid: false, message: `Mês inválido: "${monthStr}".` };
                            const startDate = new Date(Date.UTC(year, month - 1, 1));
                            const endDate = new Date(Date.UTC(year, month, 0));
                            periods.push({ start: startDate, end: endDate, monthStr: monthStr.replace('/', '-') });
                        }
                    } else {
                        const startStr = OuvidoriasApp.elements.meanDataInicio.value.trim();
                        const endStr = OuvidoriasApp.elements.meanDataFim.value.trim();
                        if (!startStr || !endStr) return { isValid: false, message: "Data de Início e Fim são obrigatórias." };
                        const dateRegex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
                        const startMatch = startStr.match(dateRegex);
                        const endMatch = endStr.match(dateRegex);
                        if (!startMatch || !endMatch) return { isValid: false, message: "Formato de data inválido. Use DD/MM/AAAA." };
                        const startDate = new Date(Date.UTC(startMatch[3], startMatch[2] - 1, startMatch[1]));
                        const endDate = new Date(Date.UTC(endMatch[3], endMatch[2] - 1, endMatch[1]));
                        if (endDate < startDate) return { isValid: false, message: "A data final não pode ser anterior à data inicial." };
                        periods.push({ start: startDate, end: endDate });
                    }
                    const ouvidoriaId = OuvidoriasApp.elements.meanOuvidoriaId.value.trim();
                    if (!ouvidoriaId) return { isValid: false, message: "'ID para Nome do Arquivo' é obrigatório." };
                    return {
                        isValid: true,
                        data: {
                            files: meanFiles,
                            codLinhas: codLinhas.toUpperCase() === 'TODAS' ? 'TODAS' : codLinhas.split(',').map(l => l.trim()),
                            empresas: empresasValidadas,
                            periods,
                            isMultiMonth,
                            specificMonthsStr,
                            tipoDia: OuvidoriasApp.elements.meanTipoDia.value,
                            processEquivalent: OuvidoriasApp.elements.meanProcessEquivalent.checked,
                            fileName: ouvidoriaId.endsWith('.xlsx') ? ouvidoriaId : `${ouvidoriaId}.xlsx`
                        }
                    };
                },

                async parseMeanTxtFile(file, requiredCols) {
                    return new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onload = (event) => {
                            try {
                                const text = event.target.result;
                                const lines = text.split(/\r\n|\n/).filter(Boolean);
                                if (lines.length < 2) return resolve([]);
                                const delimiters = [';', '\t', ','];
                                let delimiter = ';';
                                let maxCols = 0;
                                for (const d of delimiters) {
                                    const cols = lines[0].split(d).length;
                                    if (cols > maxCols) { maxCols = cols; delimiter = d; }
                                }
                                const header = lines[0].split(delimiter).map(h => h.trim().toUpperCase());
                                const requiredColsUpper = requiredCols.map(c => c.toUpperCase());
                                const colIndices = {};
                                for (const col of requiredColsUpper) {
                                    const index = header.indexOf(col);
                                    if (index === -1) {
                                        if (['NMPASSEQUIVALENTE', 'DSDIATIPO'].includes(col)) continue;
                                        throw new Error(`Coluna obrigatória não encontrada: ${col}.`);
                                    }
                                    colIndices[col] = index;
                                }
                                const data = lines.slice(1).map(line => {
                                    const parts = line.split(delimiter);
                                    if (parts.length < maxCols) return null;
                                    let row = {};
                                    for (const col in colIndices) { row[col] = parts[colIndices[col]]?.trim() || ''; }
                                    return row;
                                }).filter(Boolean);
                                resolve(data);
                            } catch (e) { reject(e); }
                        };
                        reader.onerror = () => reject(new Error(`Erro ao ler o arquivo ${file.name}`));
                        reader.readAsText(file, 'latin1');
                    });
                },

                async processMeanData(inputs, allRawData) {
                    const processedData = allRawData
                        .filter(row => inputs.periods.some(p => row.date >= p.start && row.date <= p.end))
                        .filter(row => inputs.empresas.includes(row.CDOPERADOR))
                        .filter(row => inputs.codLinhas === 'TODAS' || inputs.codLinhas.includes(row.CDLINHA))
                        .filter(row => {
                            if (inputs.tipoDia === 'TODOS') return true;
                            return row.DSDIATIPO === inputs.tipoDia;
                        });
                    const lineTotals = new Map();
                    processedData.forEach(row => {
                        const key = `${row.CDOPERADOR}|${row.CDLINHA}`;
                        const dateStr = row.date.toISOString().split('T')[0];
                        if (!lineTotals.has(key)) {
                            lineTotals.set(key, { totalPass: 0, totalEquiv: 0, uniqueDays: new Set() });
                        }
                        const totals = lineTotals.get(key);
                        totals.totalPass += row.passTotal;
                        totals.totalEquiv += row.passEquiv;
                        totals.uniqueDays.add(dateStr);
                    });
                    const excelData = [];
                    for (const [key, value] of lineTotals.entries()) {
                        const [operador, linha] = key.split('|');
                        const dayCount = value.uniqueDays.size;
                        if (dayCount > 0) {
                            excelData.push({
                                CDOPERADOR: operador,
                                CDLINHA: linha,
                                DIAS_CONTABILIZADOS: dayCount,
                                "MÉDIA_PASSAGEIROS": value.totalPass / dayCount,
                                "MÉDIA_EQUIVALENTE": inputs.processEquivalent ? (value.totalEquiv / dayCount) : undefined,
                            });
                        }
                    }
                    excelData.sort((a, b) => {
                        if (a.CDOPERADOR < b.CDOPERADOR) return -1;
                        if (a.CDOPERADOR > b.CDOPERADOR) return 1;
                        const linhaA = parseInt(a.CDLINHA, 10);
                        const linhaB = parseInt(b.CDLINHA, 10);
                        if (linhaA < linhaB) return -1;
                        if (linhaA > linhaB) return 1;
                        return 0;
                    });
                    return excelData;
                }
            }
        }
    };

    OuvidoriasApp.init();
});