document.addEventListener("DOMContentLoaded", () => {
    if (typeof XLSX === 'undefined') console.error("Biblioteca XLSX não carregada.");
    if (typeof ExcelJS === 'undefined') console.error("Biblioteca ExcelJS não carregada.");

    class Config {
        static get ORDEM_EMPRESAS() {
            return ['BOA', 'CAX', 'CSR', 'EME', 'GLO', 'SJT', 'VML'];
        }

        static get SHEET_NAMES() {
            return {
                EXTRAS: '_VIAGENS_EXTRAS',
                REDUCOES: 'REDUÇÃO_DE_SERVIÇOS'
            };
        }

        static get EP_COLUNAS_SAIDA() {
            return [
                "CDOPERADOR", "CDLINHA", "DTOPERACAO", "NMQTDVIAGENSMETA",
                "NMFROTAMETA", "NMEXTUTILMETA", "NMEXTMORTAMETA",
                "NMQTDVIAGENSREF", "NMFROTAREF", "NMEXTUTILREF",
                "NMEXTMORTAREF", "DSMOTIVO"
            ];
        }

        static get STYLES() {
            return {
                colorPrimary: 'FFD95F02',
                colorSecondary: 'FF0038A8',
                colorWhite: 'FFFFFFFF',
                colorLightGray: 'FFF0F2F5',
                colorTotalBg: 'FFFFF0E6',
                colorBorder: 'FFBFBFBF',
                fontFamily: 'Calibri'
            };
        }
    }

    class State {
        constructor() {
            this.fullExecutionList = [];
            this.epDiasSelecionados = Array.from({ length: 15 }, (_, i) => i + 1);
            this.cleanupFunctions = {
                congTxt: null,
                congXls: null,
                ajuste: null,
                epTxt: null
            };
        }
    }

    class Utils {
        static readFileAsText(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve(e.target.result);
                reader.onerror = (e) => reject(new Error(`Erro ao ler ${file.name}: ${e}`));
                reader.readAsText(file, 'latin1');
            });
        }

        static readSheetFile(file, options = {}) {
            const { sheetName, sheetIndex = 0, headerRowIndex = 0 } = options;
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
                        let targetName = sheetName;

                        if (!targetName && workbook.SheetNames.length > sheetIndex) {
                            targetName = workbook.SheetNames[sheetIndex];
                        } else if (sheetName && !workbook.Sheets[sheetName]) {
                            targetName = workbook.SheetNames.find(s => s.trim().toLowerCase() === sheetName.trim().toLowerCase());
                        }

                        if (!targetName || !workbook.Sheets[targetName]) {
                            if (sheetName === Config.SHEET_NAMES.EXTRAS || sheetName === Config.SHEET_NAMES.REDUCOES) {
                                return resolve([]);
                            }
                            return reject(new Error(`Planilha '${sheetName || sheetIndex}' não encontrada.`));
                        }

                        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[targetName], { range: headerRowIndex, defval: "" });
                        resolve(jsonData);
                    } catch (err) {
                        console.warn("Aviso de leitura XLSX/ODS:", err.message);
                        reject(new Error(`Falha ao ler planilha: ${err.message}`));
                    }
                };
                reader.readAsArrayBuffer(file);
            });
        }

        static excelSerialDateToJSDate(serial) {
            if (typeof serial !== 'number' || isNaN(serial)) return null;
            const days = serial - (serial > 59 ? 1 : 0);
            const excelEpoch = Date.UTC(1899, 11, 30);
            return new Date(excelEpoch + days * 86400000);
        }

        static parseDate(raw) {
            if (raw instanceof Date) return raw;
            if (typeof raw === 'number') return this.excelSerialDateToJSDate(raw);
            if (typeof raw === 'string') {
                const p = raw.split('/');
                if (p.length === 3) return new Date(Date.UTC(p[2], p[1] - 1, p[0]));
            }
            return null;
        }

        static sortOperatorLine(a, b) {
            const opA = String(a.Operador || a.CDOPERADOR || '').toUpperCase();
            const opB = String(b.Operador || b.CDOPERADOR || '').toUpperCase();
            const opComp = opA.localeCompare(opB);
            if (opComp !== 0) return opComp;

            const linA = parseInt(a.Linha || a.CDLINHA, 10);
            const linB = parseInt(b.Linha || b.CDLINHA, 10);
            return (!isNaN(linA) && !isNaN(linB)) ? linA - linB : String(a.Linha).localeCompare(String(b.Linha));
        }

        static async loadCTMLogo() {
            try {
                const response = await fetch('/static/images/ctm.png');
                if (!response.ok) throw new Error("Imagem não encontrada");
                const blob = await response.blob();
                return await blob.arrayBuffer();
            } catch (e) {
                console.warn("Logo CTM não carregada:", e);
                return null;
            }
        }
    }

    class ExcelService {
        constructor(uiManager) {
            this.ui = uiManager;
        }

        async generateStyledExcel(sheetsData, filename) {
            const workbook = new ExcelJS.Workbook();
            const styles = Config.STYLES;
            const todayStr = new Date().toLocaleDateString('pt-BR') + ' ' + new Date().toLocaleTimeString('pt-BR');

            const logoBuffer = await Utils.loadCTMLogo();
            let logoId = null;
            if (logoBuffer) {
                logoId = workbook.addImage({
                    buffer: logoBuffer,
                    extension: 'png',
                });
            }

            const commonStyles = {
                title: { font: { name: 'Helvetica', size: 18, bold: true, color: { argb: styles.colorSecondary } }, alignment: { horizontal: 'center', vertical: 'middle' } },
                subtitle: { font: { name: 'Helvetica', size: 12, color: { argb: 'FF505050' } }, alignment: { horizontal: 'center', vertical: 'middle' } },
                header: { font: { name: styles.fontFamily, size: 11, bold: true, color: { argb: styles.colorWhite } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: styles.colorSecondary } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin', color: { argb: styles.colorWhite } }, left: { style: 'thin', color: { argb: styles.colorWhite } }, bottom: { style: 'thin', color: { argb: styles.colorWhite } }, right: { style: 'thin', color: { argb: styles.colorWhite } } } },
                cellOdd: { font: { name: styles.fontFamily, size: 11, color: { argb: 'FF000000' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin', color: { argb: styles.colorBorder } }, bottom: { style: 'thin', color: { argb: styles.colorBorder } }, left: { style: 'thin', color: { argb: styles.colorBorder } }, right: { style: 'thin', color: { argb: styles.colorBorder } } } },
                cellEven: { font: { name: styles.fontFamily, size: 11, color: { argb: 'FF000000' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: styles.colorLightGray } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: { style: 'thin', color: { argb: styles.colorBorder } }, bottom: { style: 'thin', color: { argb: styles.colorBorder } }, left: { style: 'thin', color: { argb: styles.colorBorder } }, right: { style: 'thin', color: { argb: styles.colorBorder } } } }
            };

            // Colunas que DEVEM ter casas decimais (Log e outros)
            const logNumericCols = new Set(["Programada", "Realizada", "Limites", "Validadas", "ExtraODS", "ReducaoODS", "ValorFinal"]);

            for (const sheetConf of sheetsData) {
                const ws = workbook.addWorksheet(sheetConf.name, {
                    views: [{ state: 'frozen', ySplit: 7, showGridLines: false }]
                });

                const data = sheetConf.data;
                const columns = sheetConf.columns;
                const lastColIndex = columns.length;

                // --- Cabeçalhos ---
                ws.mergeCells(1, 1, 1, Math.max(5, lastColIndex - 4));
                ws.getCell(1, 1).value = "GRANDE RECIFE CONSÓRCIO DE TRANSPORTE";
                ws.getCell(1, 1).style = commonStyles.title;

                ws.mergeCells(2, 1, 2, Math.max(5, lastColIndex - 4));
                ws.getCell(2, 1).value = sheetConf.title || "RELATÓRIO DETALHADO";
                ws.getCell(2, 1).style = commonStyles.title;
                ws.getCell(2, 1).font = { ...commonStyles.title.font, size: 14 };

                ws.mergeCells(3, 1, 3, Math.max(5, lastColIndex - 4));
                ws.getCell(3, 1).value = "DGFC - DIVISÃO DE GESTÃO FINANCEIRA DOS CONTRATOS";
                ws.getCell(3, 1).style = commonStyles.subtitle;

                ws.mergeCells(4, 1, 4, Math.max(5, lastColIndex - 4));
                ws.getCell(4, 1).value = `GERADO EM: ${todayStr}`;
                ws.getCell(4, 1).style = commonStyles.subtitle;
                ws.getCell(4, 1).font = { size: 10, italic: true };

                if (logoId !== null) {
                    let colIndex = 7;
                    if (['Limites', 'Saldo', 'Validadas'].includes(sheetConf.name)) colIndex = 12;
                    ws.addImage(logoId, { tl: { col: colIndex, row: 0.2 }, ext: { width: 350, height: 130 }, editAs: 'absolute' });
                }

                ws.addRow([]);
                ws.addRow([]);

                const headerRow = ws.addRow(columns);
                headerRow.height = 25;
                headerRow.eachCell((cell) => { cell.style = commonStyles.header; });

                // --- PROCESSAMENTO DE DADOS ---
                data.forEach((item, index) => {
                    const rowValues = columns.map(col => {
                        let val = item[col];
                        const colName = String(col).trim();

                        // 1. TRATAMENTO PARA LINHA: Garante Inteiro Puro
                        if (colName === 'Linha') {
                            if (val === null || val === undefined || val === '') return 0;
                            // Remove virgula, converte pra float e arredonda pra baixo (inteiro)
                            const num = Math.floor(parseFloat(String(val).replace(',', '.'))); 
                            return isNaN(num) ? 0 : num;
                        }

                        // 2. TRATAMENTO PARA OUTROS: Mantém decimais
                        if (typeof val === 'string') {
                            val = val.trim();
                            if (/^-?\d+([.,]\d+)?$/.test(val)) return parseFloat(val.replace(',', '.'));
                        }
                        return val;
                    });

                    const row = ws.addRow(rowValues);
                    const isEven = (index % 2 === 0);

                    row.eachCell((cell, colNum) => {
                        cell.style = isEven ? commonStyles.cellEven : commonStyles.cellOdd;
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };

                        const colName = columns[colNum - 1];

                        // --- FORMATAÇÃO VISUAL ---
                        if (colName === 'Linha') {
                            // FORÇA BRUTA: Se é Linha, é Inteiro '0'. Sem exceção.
                            cell.numFmt = '0';
                            // Garante novamente que o valor na célula é int
                            cell.value = parseInt(cell.value); 
                        } 
                        else if (typeof cell.value === 'number') {
                            // Lógica específica para colunas que PRECISAM de decimal
                            if (sheetConf.name === "Log" && logNumericCols.has(colName)) {
                                cell.numFmt = '#,##0.0';
                            } else if (["Limites", "Saldo", "Validadas"].includes(sheetConf.name) && !isNaN(parseInt(colName))) {
                                cell.numFmt = '#,##0.0';
                            } else {
                                cell.numFmt = '#,##0';
                            }
                        }
                    });
                });

                ws.columns.forEach((column) => {
                    let maxLength = 0;
                    column.eachCell({ includeEmpty: true }, (cell) => {
                        if (cell.row < 7) return; 
                        let cellLength = 0;
                        if (cell.value) {
                            cellLength = cell.value.toString().length;
                            if (cell.value instanceof Date) cellLength = 12; 
                        }
                        if (cellLength > maxLength) maxLength = cellLength;
                    });
                    column.width = Math.min(Math.max(maxLength + 2, 10), 50);
                });
            }

            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            if ('showSaveFilePicker' in window) {
                try {
                    const handle = await window.showSaveFilePicker({
                        suggestedName: filename,
                        types: [{ description: 'Arquivo Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
                    });
                    const writable = await handle.createWritable();
                    await writable.write(blob);
                    await writable.close();
                    this.ui.showFlashMessage("Arquivo salvo com sucesso.", 'success');
                    return;
                } catch (err) { if (err.name !== 'AbortError') console.warn(err); }
            }

            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            this.ui.showFlashMessage("Arquivo baixado.", 'success');
        }
    }

    class DataService {
        async processarCongestionamento(filesTxt, filesXls, fileOds) {
            const [extrasMap, reducoesMap] = await this.carregarDadosODS(fileOds);
            const dadosTxtCalculados = await this.processarArquivosTXT(filesTxt, extrasMap, reducoesMap);
            const dadosXls = await this.processarArquivosXLS(filesXls);
            return this.consolidarDados(dadosTxtCalculados, dadosXls);
        }

        async carregarDadosODS(fileOds) {
            const processarAba = async (sheetName, rowIndex) => {
                try {
                    const raw = await Utils.readSheetFile(fileOds, { sheetName, headerRowIndex: rowIndex });
                    const map = new Map();
                    raw.forEach(row => {
                        const keys = Object.keys(row);
                        const getVal = (s) => row[keys.find(k => k.toLowerCase().includes(s))];

                        const dataObj = Utils.parseDate(getVal('data'));
                        const emp = getVal('empresa');
                        let lin = getVal('linha');
                        if (lin) {
                            const p = parseInt(String(lin).trim(), 10);
                            if (!isNaN(p)) lin = String(p);
                            else lin = String(lin).trim();
                        }
                        const val = getVal('quantidade') || getVal('horario');

                        if (!dataObj || !emp || !lin || val == null) return;

                        const dataKey = `${dataObj.getUTCDate()}/${dataObj.getUTCMonth() + 1}/${dataObj.getUTCFullYear()}`;
                        const key = `${String(emp).trim().toUpperCase()}|${String(lin).trim()}|${dataKey}`;
                        const qtd = parseFloat(String(val).replace(',', '.')) || 0;

                        map.set(key, (map.get(key) || 0) + qtd);
                    });
                    return map;
                } catch (e) {
                    console.warn(`Aviso ODS (${sheetName}): ${e.message}`);
                    return new Map();
                }
            };
            const extras = await processarAba(Config.SHEET_NAMES.EXTRAS, 8);
            const reducoes = await processarAba(Config.SHEET_NAMES.REDUCOES, 9);
            return [extras, reducoes];
        }

        async processarArquivosTXT(filesTxt, extrasMap, reducoesMap) {
            let resultados = [];
            for (const file of filesTxt) {
                const content = await Utils.readFileAsText(file);
                const lines = content.replace(/\r/g, "").split('\n');
                let headerIndex = lines.findIndex(l => l.toLowerCase().includes('operador') && l.toLowerCase().includes('linha'));
                if (headerIndex === -1) headerIndex = 1;

                const headers = lines[headerIndex].split('\t').map(h => h.trim());
                const valueCols = headers.filter(h => /^\d+$/.test(h));

                const rows = lines.slice(headerIndex + 1).map(line => {
                    const cols = line.split('\t');
                    const obj = {};
                    headers.forEach((h, i) => obj[h] = cols[i]?.trim() || '');
                    return obj;
                }).filter(r => r['Descrição'] && r['Linha']);

                const calcMap = {};
                rows.forEach(row => {
                    const desc = row['Descrição'];
                    if (desc !== "Viagem Programada" && desc !== "Viagem Realizada") return;
                    
                    const rawL = row['Linha'].split(" - ")[0].trim();
                    const pL = parseInt(rawL, 10);
                    const linhaLimpa = isNaN(pL) ? rawL : String(pL);
                    
                    const metaKey = `${row.Operador}|${linhaLimpa}|${row.Ano}|${row.Mês}|${row.Quinzena}`;

                    valueCols.forEach(dia => {
                        const val = parseFloat((row[dia] || '0').replace(',', '.')) || 0;
                        const fullKey = `${metaKey}|${dia}`;
                        if (!calcMap[fullKey]) {
                            calcMap[fullKey] = {
                                Operador: row.Operador, Linha: linhaLimpa,
                                Ano: row.Ano, Mês: row.Mês, Quinzena: row.Quinzena, Dia: dia,
                                Prog: 0, Real: 0
                            };
                        }
                        if (desc === "Viagem Programada") calcMap[fullKey].Prog = val;
                        if (desc === "Viagem Realizada") calcMap[fullKey].Real = val;
                    });
                });

                Object.values(calcMap).forEach(item => {
                    const keyLookup = `${item.Operador}|${item.Linha}|${parseInt(item.Dia)}/${parseInt(item.Mês)}/${parseInt(item.Ano)}`;
                    const extra = extrasMap.get(keyLookup) || 0;
                    const rawRed = reducoesMap.get(keyLookup) || 0;
                    const reducao = rawRed === 0 ? 0 : -Math.abs(rawRed);
                    const obs = (item.Real > item.Prog && extra > 0) ? `Extra aplicado (${extra})` : "Normal";

                    resultados.push({ ...item, Obs: obs, ExtraLog: extra, ReducaoLog: reducao });
                });
            }
            return resultados;
        }

        async processarArquivosXLS(filesXls) {
            let dados = [];
            for (const f of filesXls) {
                try {
                    const raw = await Utils.readSheetFile(f, { headerRowIndex: 11 });
                    const sigla = f.name.substring(0, 3).toUpperCase();
                    if (!Config.ORDEM_EMPRESAS.includes(sigla)) continue;

                    const limpos = raw.map(row => {
                        const kLinha = Object.keys(row).find(k => k.toLowerCase().includes('linha'));
                        if (!kLinha) return null;
                        
                        let lVal = String(row[kLinha]).trim();
                        const pLv = parseInt(lVal, 10);
                        if (!isNaN(pLv)) lVal = String(pLv);

                        const newRow = { Operador: sigla, Linha: lVal };
                        Object.keys(row).forEach(k => {
                            if (!isNaN(parseFloat(k))) {
                                const val = parseFloat(String(row[k]).replace('$', '')) || 0;
                                newRow[parseInt(k, 10)] = val === 0 ? 0 : -Math.abs(val);
                            }
                        });
                        return newRow;
                    }).filter(Boolean);
                    dados.push(...limpos);
                } catch (e) { console.warn("XLS Error:", e); }
            }
            return dados;
        }

        consolidarDados(dadosTxt, dadosXls) {
            const xlsMap = new Map();
            dadosXls.forEach(r => xlsMap.set(`${r.Operador}|${r.Linha}`, r));

            const dfFiltrado = dadosTxt.filter(item => {
                const key = `${item.Operador}|${item.Linha}`;
                const xlsRow = xlsMap.get(key);
                const valRef = xlsRow ? (xlsRow[parseInt(item.Dia)] || 0) : 0;
                
                if (item.ExtraLog === 0 && item.ReducaoLog === 0 && valRef === 0) return false;
                
                item.ValidadasLog = valRef;

                // --- NOVA LÓGICA DE CÁLCULO E OBS ---
                const prog = item.Prog;
                const real = item.Real;
                const limit = prog - real; // Ex: 85 - 105 = -20
                const extra = item.ExtraLog;
                const val = item.ValidadasLog;
                const red = item.ReducaoLog;

                let valorFinal = 0;
                let obsDetails = "";

                if (prog === 0 && real === 0 && (extra > 0 || red !== 0 || val !== 0)) {
                    valorFinal = 0;
                    obsDetails = "Linha não contém dados no arquivo de Viagem e Frota";
                } else if (extra > 0) {
                    // Regra: Se tem extra, Valor = Extra + Val + Red (lembrando que Val e Red são negativos)
                    valorFinal = extra + val + red;
                    
                    // Detalhar a conta na Obs
                    obsDetails = `Limite: ${limit.toFixed(1)} | Extra: ${extra} | Val: ${val} | Red: ${red} -> Cálculo: ${extra} + (${val}) + (${red}) = ${valorFinal.toFixed(1)}`;
                } else {
                    // Lógica original para quando NÃO tem Extra
                    // Se Prog > Real (Déficit positivo, falta viagem)
                    if (prog > real) {
                        const deficit = prog - real;
                        const potencial = val + red; // Extra é 0
                       
                        if (Math.abs(potencial) > deficit) {
                            valorFinal = -deficit;
                            obsDetails = `Limite: ${limit.toFixed(1)} | Val: ${val} | Red: ${red} -> Soma (${potencial}) > Déficit. Teto: -${deficit}`;
                        } else {
                            valorFinal = potencial;
                            obsDetails = `Limite: ${limit.toFixed(1)} | Val: ${val} | Red: ${red} -> Cobertura parcial: ${potencial}`;
                        }
                    } else {
                        // Real >= Prog E Extra == 0
                        valorFinal = 0;
                        obsDetails = `Limite: ${limit.toFixed(1)} (Real >= Prog) sem Extra. Valor: 0`;
                    }
                }

                item.ValorFinal = valorFinal;
                item.Obs = obsDetails;
                return true;
            });

            dfFiltrado.sort((a, b) => Utils.sortOperatorLine(a, b));
            dadosXls.sort((a, b) => Utils.sortOperatorLine(a, b));

            const txtPivot = {};
            dfFiltrado.forEach(d => {
                const k = `${d.Operador}|${d.Linha}|${d.Ano}|${d.Mês}|${d.Quinzena}`;
                if (!txtPivot[k]) txtPivot[k] = { Operador: d.Operador, Linha: d.Linha, Ano: d.Ano, Mês: d.Mês, Quinzena: d.Quinzena, Descrição: "Saldo" };
                txtPivot[k][d.Dia] = d.ValorFinal;
            });

            const dfTxtFinal = Object.values(txtPivot).sort((a, b) => Utils.sortOperatorLine(a, b));
            const allCols = Object.keys(dadosXls[0] || {}).filter(c => /^\d+$/.test(c)).sort((a, b) => a - b);

            const integrados = [];
            const txtMap = new Map();
            dfTxtFinal.forEach(r => txtMap.set(`${r.Operador}|${r.Linha}`, r));

            dadosXls.forEach(xlsRow => {
                const txtRow = txtMap.get(`${xlsRow.Operador}|${xlsRow.Linha}`);
                if (!txtRow) return;
                const novaLinha = { ...txtRow, Descrição: "Integrado" };
                allCols.forEach(dia => {
                    const vt = txtRow[dia] || 0;
                    const vx = xlsRow[dia] || 0;
                    novaLinha[dia] = (vt === vx || vt < vx) ? vt : vx;
                });
                integrados.push(novaLinha);
            });
            integrados.sort((a, b) => Utils.sortOperatorLine(a, b));

            return { dfFiltrado, dfTxtFinal, dadosXls, integrados, allCols };
        }

        async processarEP(filesTxt, operador, linhasStr, dias, nomeSaida) {
            if (!filesTxt?.length || !dias?.length) throw new Error("Arquivos TXT e Dias necessários.");
            const linhasSet = linhasStr.toUpperCase() === 'TODAS' ? null : new Set(linhasStr.split(',').map(s => s.trim()));
            let output = [];

            for (const file of filesTxt) {
                const content = await Utils.readFileAsText(file);
                const lines = content.replace(/\r/g, "").split('\n');
                const headerIdx = lines.findIndex(l => l.includes('Linha') && l.includes('Descrição'));
                if (headerIdx === -1) continue;

                const parts = file.name.split('-');
                const opArquivo = parts[0];
                if (operador !== 'Todas' && opArquivo !== operador) continue;

                const datePart = parts[1];
                const ano = "20" + datePart.substring(0, 2);
                const mes = datePart.substring(2, 4);
                const headers = lines[headerIdx].split('\t').map(h => h.trim());

                const rows = lines.slice(headerIdx + 1).map(l => {
                    const c = l.split('\t');
                    const o = {};
                    headers.forEach((h, i) => o[h] = c[i]?.trim());
                    return o;
                }).filter(r => r.Linha && r.Descrição);

                const grupos = {};
                rows.forEach(r => {
                    const cod = r.Linha.split(' - ')[0].trim();
                    if (linhasSet && !linhasSet.has(cod)) return;
                    if (!grupos[cod]) grupos[cod] = [];
                    grupos[cod].push(r);
                });

                for (const lin in grupos) {
                    const g = grupos[lin];
                    const prog = g.find(x => x.Descrição === "Viagem Programada");
                    const real = g.find(x => x.Descrição === "Viagem Realizada");

                    if (prog && real) {
                        dias.forEach(d => {
                            const dStr = String(d);
                            if (headers.includes(dStr)) {
                                const vp = parseFloat(prog[dStr].replace(',', '.')) || 0;
                                const vr = parseFloat(real[dStr].replace(',', '.')) || 0;
                                const diff = vr - vp;
                                if (diff < 0) {
                                    output.push({
                                        CDOPERADOR: opArquivo, CDLINHA: lin,
                                        DTOPERACAO: `${String(d).padStart(2, '0')}/${mes}/${ano}`,
                                        NMQTDVIAGENSMETA: String(diff.toFixed(2)).replace('.', ','),
                                        DSMOTIVO: "Con", NMFROTAMETA: "", NMEXTUTILMETA: "", NMEXTMORTAMETA: "",
                                        NMQTDVIAGENSREF: "", NMFROTAREF: "", NMEXTUTILREF: "", NMEXTMORTAREF: ""
                                    });
                                }
                            }
                        });
                    }
                }
            }
            if (output.length === 0) throw new Error("Nenhum dado gerado para EP.");

            output.sort((a, b) => Utils.sortOperatorLine(a, b));
            let txt = Config.EP_COLUNAS_SAIDA.join('\t') + '\r\n';
            output.forEach(r => txt += Config.EP_COLUNAS_SAIDA.map(c => r[c]).join('\t') + '\r\n');

            const blob = new Blob([txt], { type: 'text/tab-separated-values;charset=utf-8;' });
            if ('showSaveFilePicker' in window) {
                try {
                    const handle = await window.showSaveFilePicker({
                        suggestedName: nomeSaida,
                        types: [{ description: 'Arquivo Texto', accept: { 'text/plain': ['.txt'] } }]
                    });
                    const writable = await handle.createWritable();
                    await writable.write(blob);
                    await writable.close();
                } catch (e) { console.warn("Salvamento cancelado ou erro"); }
            } else {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = nomeSaida;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }
        }
    }

    class UIManager {
        constructor(state) {
            this.state = state;
            this.elements = {
                loadingIndicator: document.getElementById('loading-indicator-cong'),
                flashMessagesContainer: document.getElementById('flash-messages-container-cong'),
                executionListContainer: document.getElementById('executionListContainer'),
                refreshExecutionListButton: document.getElementById('refreshExecutionList'),
                filterEmpresaExec: document.getElementById('filterEmpresaExec'),
                filterQuinzenaExec: document.getElementById('filterQuinzenaExec'),
                filterMesExec: document.getElementById('filterMesExec'),
                filterAnoExec: document.getElementById('filterAnoExec'),
                clearExecutionFiltersButton: document.getElementById('clearExecutionFilters'),
                diasModal: document.getElementById('diasModal'),
                closeDiasModalBtn: document.getElementById('closeDiasModal'),
                openDiasModalBtn: document.getElementById('ep-dias-btn'),
                diasGridContainer: document.getElementById('dias-grid-container'),
                diasModalSalvar: document.getElementById('diasModalSalvar'),
                diasModalLimpar: document.getElementById('diasModalLimpar'),
                epDiasLabel: document.getElementById('ep-dias-label'),
                epDiasHiddenInput: document.getElementById('ep-dias-hidden-input'),
                congTxtInput: document.getElementById('cong-txt-input'),
                congXlsInput: document.getElementById('cong-xls-input'),
                ajusteAlteracoesInput: document.getElementById('ajuste-alteracoes-input'),
                congNomeSaida: document.getElementById('cong-nome-saida'),
                congLimparBtn: document.getElementById('cong-limpar-btn'),
                epTxtInput: document.getElementById('ep-txt-input'),
                epOperador: document.getElementById('ep-operador'),
                epLinhas: document.getElementById('ep-linhas'),
                epNomeSaida: document.getElementById('ep-nome-saida'),
                epLimparBtn: document.getElementById('ep-limpar-btn')
            };
        }

        initInputs() {
            const handleCongTxtChange = (files) => {
                if (files && files.length > 0) {
                    const match = files[0].name.match(/-(\d{6})-/);
                    if (match && match[1]) {
                        const outputInput = document.getElementById('cong-nome-saida');
                        if (outputInput) outputInput.value = `AlteracoesProgramacao_AJUSTADO_${match[1]}.txt`;
                    }
                }
            };
            this.state.cleanupFunctions.congTxt = this.bindFileInput('cong-txt-input', 'cong-txt-list-preview', 'cong-txt-dropzone', true, handleCongTxtChange);
            this.state.cleanupFunctions.congXls = this.bindFileInput('cong-xls-input', 'cong-xls-list-preview', 'cong-xls-dropzone', true);
            this.state.cleanupFunctions.ajuste = this.bindSingleInput('ajuste-alteracoes-input', 'ajuste-alteracoes-label');
            this.state.cleanupFunctions.epTxt = this.bindSingleInput('ep-txt-input', 'ep-txt-label');
            this.atualizarLabelDias();
        }

        setLoading(isLoading) {
            const el = this.elements.loadingIndicator;
            if (el) isLoading ? el.classList.remove('hidden') : el.classList.add('hidden');
            document.querySelectorAll('button, input, select').forEach(b => {
                b.disabled = isLoading;
                b.classList.toggle('opacity-70', isLoading);
                b.classList.toggle('cursor-not-allowed', isLoading);
            });
        }

        showFlashMessage(msg, type = 'info', duration = 7000) {
            const container = this.elements.flashMessagesContainer;
            if (!container) return;
            const icons = { success: 'check_circle', error: 'warning', info: 'info' };
            const div = document.createElement('div');
            div.className = `flash-message ${type}`;
            div.innerHTML = `<span class="material-icons-outlined mr-3">${icons[type] || 'info'}</span><span>${msg}</span><button class="ml-auto p-1.5"><span class="material-icons-outlined text-sm">close</span></button>`;
            div.querySelector('button').onclick = () => div.remove();
            container.prepend(div);
            if (duration > 0) setTimeout(() => div.remove(), duration);
            return div;
        }

        bindFileInput(inputId, previewId, dropzoneId, multiple, onFilesSelected = null) {
            const input = document.getElementById(inputId);
            const preview = document.getElementById(previewId);
            const drop = document.getElementById(dropzoneId);
            if (!input) return null;

            const update = () => {
                preview.innerHTML = '';
                if (!input.files.length) preview.innerHTML = '<span class="text-gray-400 italic">Nenhum arquivo.</span>';
                else Array.from(input.files).forEach(f => {
                    const d = document.createElement('div');
                    d.className = 'truncate';
                    d.textContent = `• ${f.name}`;
                    preview.appendChild(d);
                });
                if (onFilesSelected) onFilesSelected(input.files);
            };

            input.addEventListener('change', update);
            if (drop) {
                drop.addEventListener('click', () => input.click());
                drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('bg-gray-100'); });
                drop.addEventListener('dragleave', () => drop.classList.remove('bg-gray-100'));
                drop.addEventListener('drop', e => {
                    e.preventDefault();
                    drop.classList.remove('bg-gray-100');
                    input.files = e.dataTransfer.files;
                    update();
                });
            }
            return () => { input.value = ''; update(); };
        }

        bindSingleInput(inputId, labelId) {
            const input = document.getElementById(inputId);
            const label = document.getElementById(labelId);
            if (!input) return null;
            const defText = label.getAttribute('data-default') || 'Nenhum arquivo';
            const update = () => {
                if (input.files.length) {
                    label.textContent = input.files[0].name;
                    label.classList.add('text-blue-600', 'font-medium');
                } else {
                    label.textContent = defText;
                    label.classList.remove('text-blue-600', 'font-medium');
                }
            };
            input.addEventListener('change', update);
            return () => { input.value = ''; update(); };
        }

        atualizarLabelDias() {
            const l = this.elements.epDiasLabel;
            if (!l) return;
            const count = this.state.epDiasSelecionados.length;
            l.textContent = count === 0 ? 'Nenhum dia' : (count === 15 ? '15 dias (Padrão)' : `${count} dia(s)`);
            if (this.elements.epDiasHiddenInput) this.elements.epDiasHiddenInput.value = this.state.epDiasSelecionados.join(',');
        }

        openDiasModal() {
            const container = this.elements.diasGridContainer;
            container.innerHTML = '';
            for (let i = 1; i <= 31; i++) {
                const checked = this.state.epDiasSelecionados.includes(i) ? 'checked' : '';
                container.innerHTML += `
                    <label class="flex flex-col items-center p-2 border rounded cursor-pointer hover:bg-gray-100">
                        <input type="checkbox" data-dia="${i}" ${checked} class="mb-1 text-orange-500">
                        <span class="text-sm font-semibold">${i}</span>
                    </label>`;
            }
            this.elements.diasModal.style.display = 'flex';
            document.body.classList.add('modal-open');
        }

        closeDiasModal() {
            this.elements.diasModal.style.display = 'none';
            document.body.classList.remove('modal-open');
        }

        renderExecutionList(list) {
            const container = this.elements.executionListContainer;
            container.innerHTML = '';
            if (!list.length) {
                container.innerHTML = '<div class="execution-history-empty">Nenhum histórico.</div>';
                return;
            }
            list.forEach(id => {
                const d = document.createElement('div');
                d.className = 'execution-history-item';
                d.textContent = id;
                container.appendChild(d);
            });
        }
    }

    class AppController {
        constructor() {
            this.state = new State();
            this.ui = new UIManager(this.state);
            this.excelService = new ExcelService(this.ui);
            this.dataService = new DataService();

            this.init();
        }

        init() {
            this.ui.initInputs();
            this.bindEvents();
            this.exposeGlobalFunctions();
        }

        exposeGlobalFunctions() {
            window.iniciarProcessamentoCongestionamento = this.handleCongestionamentoSubmit.bind(this);
            window.iniciarProcessamentoEP = this.handleEPSubmit.bind(this);
            window.openDiasModal = this.ui.openDiasModal.bind(this.ui);
            
            window.switchTabCongestionamento = function(btn, targetId) {
                // Hide all contents
                document.querySelectorAll('#main-tab-content-cong .tab-content').forEach(el => el.classList.add('hidden'));
                // Show target
                document.getElementById(targetId).classList.remove('hidden');
                
                // Update buttons
                document.querySelectorAll('#main-tab button').forEach(b => {
                    b.classList.remove('border-orange-500', 'text-orange-600', 'dark:text-orange-500');
                    b.classList.add('border-transparent');
                });
                btn.classList.remove('border-transparent');
                btn.classList.add('border-orange-500', 'text-orange-600', 'dark:text-orange-500');
            };
        }

        bindEvents() {
            const el = this.ui.elements;
            el.refreshExecutionListButton?.addEventListener('click', () => this.loadHistory());
            el.clearExecutionFiltersButton?.addEventListener('click', () => {
                el.filterEmpresaExec.value = '';
                this.filterHistory();
            });
            ['filterEmpresaExec', 'filterQuinzenaExec', 'filterMesExec', 'filterAnoExec'].forEach(id => {
                el[id]?.addEventListener('input', () => setTimeout(() => this.filterHistory(), 300));
            });

            el.diasModalSalvar?.addEventListener('click', () => {
                const checks = el.diasGridContainer.querySelectorAll('input:checked');
                this.state.epDiasSelecionados = Array.from(checks).map(c => parseInt(c.dataset.dia));
                this.ui.atualizarLabelDias();
                this.ui.closeDiasModal();
            });

            el.diasModalLimpar?.addEventListener('click', () => {
                el.diasGridContainer.querySelectorAll('input').forEach(c => c.checked = false);
            });

            el.closeDiasModalBtn?.addEventListener('click', () => this.ui.closeDiasModal());

            el.congLimparBtn?.addEventListener('click', () => {
                Object.values(this.state.cleanupFunctions).forEach(fn => fn && fn());
                el.congNomeSaida.value = `AlteracoesProgramacao_AJUSTADO_${new Date().getDate()}.txt`;
                this.ui.showFlashMessage("Formulário limpo.", 'info');
            });

            el.epLimparBtn?.addEventListener('click', () => {
                if (this.state.cleanupFunctions.epTxt) this.state.cleanupFunctions.epTxt();
                this.state.epDiasSelecionados = Array.from({ length: 15 }, (_, i) => i + 1);
                this.ui.atualizarLabelDias();
                el.epOperador.value = 'Todas';
                el.epLinhas.value = 'Todas';
                this.ui.showFlashMessage("Formulário EP limpo.", 'info');
            });
        }

        loadHistory() {
            this.ui.elements.executionListContainer.innerHTML = '<div class="text-center p-4">Carregando...</div>';
            setTimeout(() => {
                this.state.fullExecutionList = [];
                this.filterHistory();
            }, 500);
        }

        filterHistory() {
            const el = this.ui.elements;
            const filters = {
                emp: el.filterEmpresaExec.value.trim().toUpperCase(),
                quin: el.filterQuinzenaExec.value,
                mes: el.filterMesExec.value,
                ano: el.filterAnoExec.value.trim()
            };

            const filtered = this.state.fullExecutionList.filter(id => {
                const m = id.match(/(\d{4})[_|-]?(\d{2})[_|-]?(\d{2})/);
                const empM = id.match(/^([A-Z]{2,4})[_|-]/i);
                if (!m) return false;
                const [_, y, mo, d] = m;
                const emp = empM ? empM[1].toUpperCase() : '';
                const q = parseInt(d) <= 15 ? '01' : '02';
                return (!filters.emp || emp.includes(filters.emp)) &&
                    (!filters.quin || q === filters.quin) &&
                    (!filters.mes || mo === filters.mes) &&
                    (!filters.ano || y === filters.ano);
            });
            this.ui.renderExecutionList(filtered);
        }

        async handleCongestionamentoSubmit(e) {
            e.preventDefault();
            this.ui.setLoading(true);
            const msg = this.ui.showFlashMessage("Processando...", 'info', 0);
            try {
                const el = this.ui.elements;
                const data = await this.dataService.processarCongestionamento(
                    Array.from(el.congTxtInput.files),
                    Array.from(el.congXlsInput.files),
                    el.ajusteAlteracoesInput.files[0]
                );

                const logData = data.dfFiltrado.map(d => ({
                    Operador: d.Operador, Linha: d.Linha, Data: `${d.Dia}/${d.Mês}/${d.Ano}`,
                    Programada: d.Prog, Realizada: d.Real, Limites: (d.Prog - d.Real),
                    Validadas: d.ValidadasLog, ExtraODS: d.ExtraLog, ReducaoODS: d.ReducaoLog,
                    ValorFinal: d.ValorFinal, Obs: d.Obs
                }));

                const colunasDinamicas = ['Operador', 'Linha', ...data.allCols];
                const cleanData = (lista) => lista.map(item => {
                    const novo = {};
                    colunasDinamicas.forEach(c => { if (item.hasOwnProperty(c)) novo[c] = item[c]; });
                    return novo;
                });

                const dadosLimites = cleanData(data.integrados);
                const dadosSaldo = cleanData(data.dfTxtFinal);
                const sheetsConfiguration = [
                    { name: "Log", title: "REGISTRO DE OPERAÇÕES (LOG)", columns: Object.keys(logData[0] || {}), data: logData },
                    { name: "Limites", title: "LIMITES OPERACIONAIS", columns: colunasDinamicas, data: dadosLimites },
                    { name: "Saldo", title: "SALDO FINAL DIÁRIO", columns: colunasDinamicas, data: dadosSaldo },
                    { name: "Validadas", title: "VIAGENS VALIDADAS (REF)", columns: colunasDinamicas, data: data.dadosXls }
                ];

                const now = new Date();
                const ts = `${now.getDate()}${now.getMonth() + 1}_${now.getHours()}${now.getMinutes()}`;
                await this.excelService.generateStyledExcel(sheetsConfiguration, `log_detalhado_${ts}.xlsx`);

                const txtOutput = [];
                data.dfFiltrado.forEach(r => {
                    if (r.ValorFinal !== 0) {
                        txtOutput.push({
                            CDOPERADOR: r.Operador, CDLINHA: r.Linha,
                            DTOPERACAO: `${String(r.Dia).padStart(2, '0')}/${String(r.Mês).padStart(2, '0')}/${r.Ano}`,
                            NMQTDVIAGENSMETA: String(r.ValorFinal.toFixed(1)).replace('.', ','),
                            DSMOTIVO: "Con", NMFROTAMETA: "", NMEXTUTILMETA: "", NMEXTMORTAMETA: "",
                            NMQTDVIAGENSREF: "", NMFROTAREF: "", NMEXTUTILREF: "", NMEXTMORTAREF: ""
                        });
                    }
                });

                txtOutput.sort((a, b) => Utils.sortOperatorLine(a, b));
                let txtContent = Config.EP_COLUNAS_SAIDA.join('\t') + '\r\n';
                txtOutput.forEach(r => txtContent += Config.EP_COLUNAS_SAIDA.map(c => r[c] || "").join('\t') + '\r\n');

                const nomeTxt = el.congNomeSaida.value || `AlteracoesProgramacao_AJUSTADO_${ts}.txt`;
                const blobTxt = new Blob([txtContent], { type: 'text/tab-separated-values;charset=utf-8;' });

                if ('showSaveFilePicker' in window) {
                    try {
                        const handle = await window.showSaveFilePicker({
                            suggestedName: nomeTxt,
                            types: [{ description: 'Arquivo Texto', accept: { 'text/plain': ['.txt'] } }]
                        });
                        const writable = await handle.createWritable();
                        await writable.write(blobTxt);
                        await writable.close();
                    } catch (e) { console.warn("Salvamento de TXT cancelado ou erro"); }
                } else {
                    const urlTxt = URL.createObjectURL(blobTxt);
                    const a = document.createElement('a');
                    a.href = urlTxt;
                    a.download = nomeTxt;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                }

            } catch (err) {
                this.ui.showFlashMessage("Erro: " + err.message, 'error', 0);
                console.error(err);
            } finally {
                this.ui.setLoading(false);
                if (msg) msg.remove();
            }
        }

        async handleEPSubmit(e) {
            e.preventDefault();
            this.ui.setLoading(true);
            const msg = this.ui.showFlashMessage("Gerando EP...", 'info', 0);
            try {
                const el = this.ui.elements;
                await this.dataService.processarEP(
                    Array.from(el.epTxtInput.files),
                    el.epOperador.value,
                    el.epLinhas.value,
                    this.state.epDiasSelecionados,
                    el.epNomeSaida.value
                );
                this.ui.showFlashMessage("EP gerado!", 'success');
            } catch (err) {
                this.ui.showFlashMessage("Erro: " + err.message, 'error', 0);
            } finally {
                this.ui.setLoading(false);
                if (msg) msg.remove();
            }
        }
    }

    new AppController();
});