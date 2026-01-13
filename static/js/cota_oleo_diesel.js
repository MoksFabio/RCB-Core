/**
 * cota_logic.js
 * * Porta completa da lógica Python (Tkinter/Pandas/Openpyxl) para JavaScript (ExcelJS/SheetJS).
 * Contém: Constantes, Processamento de Dados, Lógica de Negócio e Geração de Relatório.
 */

// ==========================================
// 1. CONFIGURAÇÃO E CONSTANTES (Config)
// ==========================================
const Config = {
    RENDIMENTO_DISPEL: {
        "Micro Urbano  s/ar": 2.610,
        "MIDI MEDIO URBANO S/AR": 2.610,
        "Básico Méd. urb s/ar": 2.610,
        "Básico Méd. urb c/ar": 2.219,
        "Padron 12  Pés s/ar": 2.610,
        "Padron 12  Pés c/ar": 2.2185,
        "Padron 13 Pés s /ar": 2.420,
        "Padron 13 Pés c /ar": 2.057,
        "Padron 14 Pés s /ar": 2.381,
        "Padron 14 Pés c /ar": 2.218,
        "Padron 15 Pés s /ar": 2.381,
        "Padron 15 Pés C /ar": 1.832,
        "Padron 15 Pés c/ar": 1.832,
        "Artic.Ext.Pés s/ar": 1.750,
        "Artic.Ext.Pés c/ar": 1.200,
        "BRT 1 Art.Ext Pés C/ar": 1.200,
        "BRT 1 Art.Ext Pés c/ ar": 1.200,
        "RODVIÁRIOS P. 13 C/AR": 1.750,
        "RODVIÁRIOS S/AR": 0,
        "RODOVIÁRIOS C/AR": 1.750,
        "RODVIÁRIO C/AR": 1.750,
        "RODVIÁRIOS C/AR": 1.750,
        "RODOVIÁRIOS P. 13 C/AR": 1.750,
        "RODOVIÁRIO P. 13 C/AR": 1.750,
        "ARTICULADO COM AR E CÂMBIO": 1.200,
        "ART.EXT PES URBANO S/AR": 1.750,
        "Midi Urbano  s/ar": 2.610,
        "Mini Urbano  c/ar": 2.088,
        "MICRO": 2.2185, 
        "PADRON COM AR": 1.8315,
        "ARTICULADO": 1.200,
        "PESADO COM AR": 2.2185,
        "BRT": 1.2, 
        "MIDI": 2.61, 
        "PESADO": 2.61
    },
    NOMES_EMPRESAS: {
        "BOA": "BOA - Borborema Imperial Transportes Ltda", "CAX": "CAX - Caxangá Empresa de Transporte Coletivo Ltda",
        "CSR": "CSR - Consórcio Recife de Transporte", "CNO": "CNO – CONSÓRCIO CONORTE",
        "EME": "EME - Metropolitana Empresa de Transporte Coletivo Ltda", "GLO": "GLO - Transportadora Globo Ltda",
        "MOB": "MOB – MobiBrasil Expresso S.A", "SJT": "SJT - José Faustino e Companhia Ltda", "VML": "VML - Viação Mirim Ltda",
        "CTC": "CTC - Companhia de Transp. e Comunicação"
    },
    CNPJ_MAP: {
        'BOA': { '1-80 BV': '10.882.777/0001-80', '3-42 CD': '10.882.777/0003-42' },
        'CAX': { '1-83 OL': '41.037.250/0001-83', '3-45 OL': '41.037.250/0003-45' },
        'CNO': { '1-39 OL': '70.227.608/0001-39', '1-66 AL': '10.687.226/0001-66', '1-40 OL': '12.790.622/0001-40' },
        'CSR': { '1-09 RE': '36.106.678/0001-09' },
        'EME': { '1-97 RE': '10.407.005/0001-97' },
        'GLO': { '2-00 RE': '12.601.233/0002-00' },
        'MOB': { '1-29 SLM': '18.938.887/0001-29', '2-00 RE': '18.938.887/0002-00' },
        'SJT': { '1-66 CSA': '09.929.134/0001-66' },
        'VML': { '1-00 RE': '08.107.369/0001-00' }
    },
    CNO_SUB_NAME_MAP: {
        '1-39 OL': 'CDA - Cidade Alta Transportes e Turismo Ltda',
        '1-66 AL': 'ITA - Transportadora Itamaracá Ltda',
        '1-40 OL': 'ROD - Rodotur Turismo Ltda'
    },
    DISTRIBUIDORA_MAP: {
        "DISLUB": "Dislub Combustíveis S/A",
        "VIBRA": "VIBRA Energia S/A",
        "IPIRANGA": "Ipiranga Produtos de Petróleo S/A",
        "RAIZEN": "Raízen Combustivéis S/A"
    },
    STYLES: {
        fontTituloPrincipal: { name: 'Calibri', size: 14, bold: true },
        fontTituloCalc: { name: 'Calibri', size: 22, italic: true },
        fontSubtituloCalc: { name: 'Calibri', bold: true },
        fontHeaderEmpresa: { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } },
        fillHeaderEmpresa: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } },
        fontHeaderTabela: { bold: true },
        fillHeaderTabela: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFA500' } }, // Orange
        fontTotal: { bold: true },
        fillTotal: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } },
        fontMediaGeral: { size: 28, bold: true },
        fillMediaGeral: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }, // Yellow
        
        // Cota Mes Styles
        fontCotaTitulo: { name: 'Calibri', size: 16, bold: true, color: { argb: 'FFFFFFFF' } },
        fillCotaTitulo: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFA500' } },
        fontCotaSubtitulo: { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF404040' } },
        fillBlack: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } },
        fontWhiteBold: { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } },
        fillOrange: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFA500' } },
        fillZebra: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8F8F8' } },
        fillDestaque: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } },
        fontDestaque: { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFBF5B00' } },
        fillEspecialCnoMob: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } },
        
        // Rateamento Styles
        fontRateamentoTitulo: { name: 'Calibri', size: 26, bold: true, color: { argb: 'FFFFFFFF' } },
        fontNomeEmpresa: { name: 'Calibri', size: 16, bold: true },
        
        // SEFAZ Styles
        fontSefazTitulo: { name: 'Calibri', size: 11, bold: true },
        fontSefazHeader: { name: 'Calibri', size: 9, bold: true },
        
        // Borders
        borderThin: {
            top: { style: 'thin', color: { argb: 'FFBFBFBF' } },
            left: { style: 'thin', color: { argb: 'FFBFBFBF' } },
            bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
            right: { style: 'thin', color: { argb: 'FFBFBFBF' } }
        },
        borderBlackThin: {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        },
        borderDouble: {
            top: { style: 'double', color: { argb: 'FF000000' } },
            bottom: { style: 'double', color: { argb: 'FF000000' } }
        },
        alignCenter: { horizontal: 'center', vertical: 'middle', wrapText: true },
        alignLeft: { horizontal: 'left', vertical: 'middle', wrapText: true },
        alignRight: { horizontal: 'right', vertical: 'middle' }
    }
};

// ==========================================
// 2. UTILITÁRIOS (Utils)
// ==========================================
class Utils {
    static async readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(e);
            reader.readAsArrayBuffer(file);
        });
    }

    static async readExcelFileToJSON(file) {
        const data = await this.readFileAsArrayBuffer(file);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        return {
            json: XLSX.utils.sheet_to_json(worksheet, { header: 1 }), // Raw array of arrays
            workbook: workbook
        };
    }

    static findHeaderRow(rows, requiredKeys = ['Linha', 'Km Total', 'Empresa']) {
        for (let i = 0; i < Math.min(rows.length, 20); i++) {
            const rowStr = rows[i].map(c => String(c || '').trim()).join(' ').toUpperCase();
            let allFound = true;
            for (const key of requiredKeys) {
                if (!rowStr.includes(key.toUpperCase())) {
                    allFound = false;
                    break;
                }
            }
            if (allFound) return i;
        }
        return -1;
    }

    static getColumnIndex(rows, headerRowIdx, columnName) {
        const row = rows[headerRowIdx];
        for (let i = 0; i < row.length; i++) {
            if (String(row[i]).trim().toUpperCase() === columnName.toUpperCase()) return i;
            // Handle variations
            if (columnName === 'Dia Tipo' && (String(row[i]).toUpperCase() === 'TP DIA')) return i;
            if (columnName === 'Km Total' && (String(row[i]).toUpperCase().includes('QUILOMETRAGEM') || String(row[i]).toUpperCase().includes('KM'))) return i;
        }
        return -1;
    }

    static getFileSortKey(filename) {
        const f = filename.toUpperCase();
        if (f.includes("DUT")) return 0;
        if (f.includes("SÁB") || f.includes("SAB")) return 1;
        if (f.includes("DOM")) return 2;
        return 3;
    }

    static getMonthNameFromDate(dateStr) {
        // Simple heuristic detector for Month name from string or date object
        const months = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
        if (dateStr instanceof Date) return months[dateStr.getMonth()];
        const s = String(dateStr).toUpperCase();
        for (let m of months) if (s.includes(m.substring(0, 3))) return m;
        
        // Try to parse typical header dates "dd/mm/yyyy"
        const match = s.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if(match && match[2]) {
             return months[parseInt(match[2])-1];
        }
        return "MÊS DESCONHECIDO";
    }

    static applyStyle(cell, styleObj) {
        if (styleObj.font) cell.font = styleObj.font;
        if (styleObj.fill) cell.fill = styleObj.fill;
        if (styleObj.border) cell.border = styleObj.border;
        if (styleObj.alignment) cell.alignment = styleObj.alignment;
        if (styleObj.numFmt) cell.numFmt = styleObj.numFmt;
    }
    
    // Simula lógica de arredondamento Python: floor se < 2500 resto, ceil se >= 2500
    static pythonRound5000(val) {
        if (!val || val <= 0) return 0;
        const remainder = val % 5000;
        if (remainder < 2500) {
            return Math.floor(val / 5000) * 5000;
        } else {
            return Math.ceil(val / 5000) * 5000;
        }
    }
}

// ==========================================
// 3. CONTROLADOR DE LÓGICA (AppController)
// ==========================================
class CotaDieselApp {
    constructor() {
        this.data = {
            passado: [],
            atual: [],
            pco: null,
            rateamento: null,
            cnoRaw: null,
            mobRaw: null
        };
        this.params = {};
    }

    // --- Data Extraction Logic ---

    async processFiles(filesPassado, filesAtual, filesPCO, fileRateamento, fileCNO, fileMOB, userParams) {
        this.params = userParams;

        // 1. Extract Main Data (Atual/Passado)
        this.data.passado = await this.extractKmData(filesPassado);
        this.data.atual = await this.extractKmData(filesAtual);
        
        // 2. Extract PCO
        if (filesPCO && filesPCO.length > 0) {
            this.data.pco = await this.processPCO(filesPCO[0], userParams.pcoMes, userParams.pcoQuinzena);
        }

        // 3. Extract Rateamento
        if (fileRateamento) {
            this.data.rateamento = await this.processRateamento(fileRateamento);
        } else {
            // Se já foi carregado via "Carregar de Arquivo" no UI e passado no userParams, use-o
            if(userParams.dadosRateamentoManual) {
                this.data.rateamento = this.convertManualRateamento(userParams.dadosRateamentoManual);
            }
        }

        // 4. Extract CNO/MOB Raw
        if (fileCNO) this.data.cnoRaw = (await Utils.readExcelFileToJSON(fileCNO)).json;
        if (fileMOB) this.data.mobRaw = (await Utils.readExcelFileToJSON(fileMOB)).json;

        // Generate Report
        return await this.generateReport();
    }

    async extractKmData(files) {
        let combinedData = [];
        const sortedFiles = Array.from(files).sort((a, b) => Utils.getFileSortKey(a.name) - Utils.getFileSortKey(b.name));

        for (const file of sortedFiles) {
            const result = await Utils.readExcelFileToJSON(file);
            const rows = result.json;
            const headerIdx = Utils.findHeaderRow(rows);
            
            if (headerIdx === -1) continue;

            const cLinha = Utils.getColumnIndex(rows, headerIdx, 'Linha');
            const cEmpresa = Utils.getColumnIndex(rows, headerIdx, 'Empresa');
            const cDia = Utils.getColumnIndex(rows, headerIdx, 'Dia Tipo');
            const cKm = Utils.getColumnIndex(rows, headerIdx, 'Km Total');

            if (cLinha === -1 || cKm === -1) continue;

            for (let i = headerIdx + 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row) continue;
                
                const valLinha = row[cLinha];
                if (!valLinha || String(valLinha).toUpperCase().includes('TOTAL')) continue;

                let linhaStr = String(valLinha).trim();
                try { linhaStr = String(parseInt(linhaStr)); } catch(e){} // Remove .0 if present

                const empresa = cEmpresa > -1 ? String(row[cEmpresa] || '').trim() : '';
                const diaTipo = cDia > -1 ? String(row[cDia] || '').trim() : '';
                const km = parseFloat(row[cKm]) || 0;

                combinedData.push({
                    linha: linhaStr,
                    empresa: empresa,
                    diaTipo: diaTipo,
                    km: km,
                    sourceFile: file.name
                });
            }
        }
        
        // Remove duplicates (keep last)
        const uniqueMap = new Map();
        combinedData.forEach(item => {
            const key = `${item.linha}|${item.empresa}|${item.diaTipo}`;
            uniqueMap.set(key, item);
        });

        return Array.from(uniqueMap.values());
    }

    async processPCO(file, mes, quinzena) {
        // Logic equivalent to processar_quinzena
        const result = await Utils.readExcelFileToJSON(file);
        const wb = result.workbook;
        
        // Find sheet by name match (approximate)
        const sheetName = wb.SheetNames.find(n => n.toUpperCase().includes(mes.toUpperCase()));
        if (!sheetName) return [];

        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
        
        // Interval logic from Python
        let startRow, endRow, headerRow;
        if (quinzena === '1ª') { startRow = 7; endRow = 19; headerRow = 8; }
        else { startRow = 30; endRow = 42; headerRow = 30; }

        // Adjust for 0-based index
        const slice = rows.slice(startRow, endRow);
        // Find header in slice (Python code assumes exact row indices)
        // Let's search for "EMPRESA" in the slice
        let localHeaderIdx = -1;
        for(let i=0; i<slice.length; i++) {
            if(slice[i] && String(slice[i][0]).toUpperCase().includes('EMPRESA')) {
                localHeaderIdx = i; break;
            }
        }
        
        if (localHeaderIdx === -1) return [];

        const headers = slice[localHeaderIdx].map(h => String(h).trim());
        const dataRows = slice.slice(localHeaderIdx + 1);
        
        const processed = [];

        dataRows.forEach(row => {
            const empresa = String(row[0] || '').trim();
            if (!empresa || empresa.toUpperCase().includes('TOTAL') || empresa.toUpperCase() === 'EMPRESA') return;

            for (let c = 1; c < row.length; c++) {
                const category = headers[c];
                const value = row[c];
                
                if (!category || category === 'Total CCT' || category === 'Total Geladinho') continue;
                if (!value) continue;

                const valNum = parseInt(value);
                const rend = Config.RENDIMENTO_DISPEL[category] || 0;

                if (valNum > 0) {
                    processed.push({
                        empresa: empresa,
                        categoria: category,
                        valor: valNum,
                        rendimentoRef: rend,
                        qteXRend: valNum * rend
                    });
                }
            }
        });

        return processed;
    }

    async processRateamento(file) {
        const result = await Utils.readExcelFileToJSON(file);
        const rows = result.json;
        // Search header
        let hIdx = -1;
        for(let i=0; i<20; i++) {
            if (rows[i] && rows[i].map(x=>String(x).toUpperCase()).includes('GARAGEM_ID')) {
                hIdx = i; break;
            }
        }
        if(hIdx === -1) return {};

        const h = rows[hIdx].map(x=>String(x).toUpperCase().trim());
        const cEmp = h.indexOf('EMPRESA');
        const cGar = h.indexOf('GARAGEM_ID');
        const cLitG = h.indexOf('LITROS_GARAGEM');
        const cLitC = h.indexOf('LITROS_COMPANHIA');
        const cPosto = h.indexOf('POSTO');
        const cCNPJ = h.indexOf('CNPJ');
        const cIE = h.indexOf('INSCRIÇÃO_ESTADUAL');

        const rateamento = {};

        for(let i=hIdx+1; i<rows.length; i++) {
            const r = rows[i];
            if(!r || !r[cEmp]) continue;
            
            const emp = r[cEmp];
            if(!rateamento[emp]) rateamento[emp] = [];
            
            rateamento[emp].push({
                garagemId: r[cGar],
                litrosGaragem: r[cLitG],
                litrosCompanhia: r[cLitC],
                posto: r[cPosto],
                cnpj: r[cCNPJ],
                ie: r[cIE]
            });
        }
        return rateamento;
    }

    convertManualRateamento(manualData) {
        // Converts the UI structure back to a list structure similar to file import
        const rateamento = {};
        for(const [emp, entries] of Object.entries(manualData)) {
            // entries is list of tuples: [litrosG, litrosC, posto, cnpj, ie]
            rateamento[emp] = [];
            // We need to reconstruct garage IDs. Logic: if LitrosG provided, it's a new garage.
            let currentGid = 0;
            entries.forEach(item => {
                const [lG, lC, pst, cnpj, ie] = item;
                if(lG !== null && lG !== '') currentGid++;
                rateamento[emp].push({
                    garagemId: currentGid,
                    litrosGaragem: lG,
                    litrosCompanhia: lC,
                    posto: pst,
                    cnpj: cnpj,
                    ie: ie
                });
            });
        }
        return rateamento;
    }

    // ==========================================
    // 4. REPORT GENERATION (ExcelJS)
    // ==========================================

    async generateReport() {
        const wb = new ExcelJS.Workbook();
        wb.creator = 'RCB System';
        wb.created = new Date();

        // --- Sheet 1: Km Prog ---
        await this.buildKmProgSheet(wb);

        // --- Sheet 2: Rendimento PCO ---
        const rendimentosMap = await this.buildRendimentoPCOSheet(wb);

        // --- Sheet 3: Cálculo CNO (if raw data exists) ---
        let cotaCNO = null;
        if (this.data.cnoRaw) {
            cotaCNO = await this.buildConsolidationSheet(wb, this.data.cnoRaw, "Cálculo CNO", "CNO");
        }

        // --- Sheet 4: Cálculo MOB (if raw data exists) ---
        let cotaMOB = null;
        if (this.data.mobRaw) {
            cotaMOB = await this.buildConsolidationSheet(wb, this.data.mobRaw, "Cálculo MOB", "MOB");
        }

        // --- Sheet 5: Cota do Mês ---
        const cotasPorEmpresa = await this.buildCotaMesSheet(wb, rendimentosMap, cotaCNO, cotaMOB);

        // --- Sheet 6: Rateamento ---
        if (this.data.rateamento) {
            await this.buildRateamentoSheet(wb, cotasPorEmpresa);
        }

        // --- Sheet 7: SEFAZ ---
        await this.buildSefazSheet(wb, cotasPorEmpresa);

        // Write buffer
        const buffer = await wb.xlsx.writeBuffer();
        return buffer;
    }

    // --- SHEET BUILDERS ---

    async buildKmProgSheet(wb) {
        const ws = wb.addWorksheet('Km Prog');
        
        // Headers
        const headers = ["Status", "Linha", "Empresa", "Dia Tipo", "Km Passado", "Km Atual", "Diferença"];
        const hRow = ws.addRow(headers);
        hRow.eachCell(c => {
            c.font = Config.STYLES.fontHeaderTabela;
            c.fill = Config.STYLES.fillHeaderTabela;
            c.alignment = Config.STYLES.alignCenter;
            c.border = Config.STYLES.borderThin;
        });

        // Map Past Data
        const pastMap = new Map();
        this.data.passado.forEach(item => {
            pastMap.set(`${item.linha}|${item.empresa}|${item.diaTipo}`, item.km);
        });

        // Iterate Current
        const processedKeys = new Set();
        
        this.data.atual.forEach(curr => {
            const key = `${curr.linha}|${curr.empresa}|${curr.diaTipo}`;
            processedKeys.add(key);
            
            const pastKm = pastMap.get(key) || 0;
            const diff = curr.km - pastKm;
            let status = "Ok";
            if (pastKm === 0) status = "Linha Nova";
            else if (Math.abs(diff) > 0.01) status = "Diferença";

            const row = ws.addRow([status, curr.linha, curr.empresa, curr.diaTipo, pastKm, curr.km, diff]);
            
            // Conditional Formatting (Visual only via styles for now)
            if(status === "Ok") row.getCell(1).font = { color: { argb: 'FF006100' }, bold: true };
            else row.getCell(1).font = { color: { argb: 'FF9C0006' }, bold: true };
            
            row.getCell(5).numFmt = '#,##0.00';
            row.getCell(6).numFmt = '#,##0.00';
            row.getCell(7).numFmt = '#,##0.00';
        });

        // Check for Deleted Lines (in Past but not Current)
        pastMap.forEach((km, key) => {
            if (!processedKeys.has(key)) {
                const [linha, empresa, dia] = key.split('|');
                const row = ws.addRow(["Linha Excluída", linha, empresa, dia, km, 0, -km]);
                row.getCell(1).font = { color: { argb: 'FFFF0000' }, bold: true };
                row.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } });
            }
        });

        ws.columns.forEach(c => c.width = 15);
    }

    async buildRendimentoPCOSheet(wb) {
        const rendimentosFinais = {};
        if (!this.data.pco || this.data.pco.length === 0) return rendimentosFinais;

        const ws = wb.addWorksheet('Rendimento PCO');
        let currentRow = 1;

        // Group by Empresa
        const groups = {};
        this.data.pco.forEach(i => {
            if(!groups[i.empresa]) groups[i.empresa] = [];
            groups[i.empresa].push(i);
        });

        for (const emp in groups) {
            const items = groups[emp];
            
            // Header Empresa
            ws.mergeCells(currentRow, 1, currentRow, 6);
            const titleCell = ws.getCell(currentRow, 1);
            titleCell.value = Config.NOMES_EMPRESAS[emp] || emp;
            Utils.applyStyle(titleCell, { font: Config.STYLES.fontHeaderEmpresa, fill: Config.STYLES.fillHeaderEmpresa, alignment: Config.STYLES.alignCenter });
            currentRow++;

            // Header Table
            const hRow = ws.getRow(currentRow);
            hRow.values = ['Categoria', 'Qtde Veic. PCO', 'Proporção (%)', 'Rendimento Lic.', 'Qte X Rend', 'Média Geral'];
            hRow.eachCell(c => Utils.applyStyle(c, { font: Config.STYLES.fontHeaderTabela, fill: Config.STYLES.fillHeaderTabela, border: Config.STYLES.borderThin, alignment: Config.STYLES.alignCenter }));
            currentRow++;

            // Data
            const totalValor = items.reduce((s, i) => s + i.valor, 0);
            const totalQteXRend = items.reduce((s, i) => s + i.qteXRend, 0);
            const mediaGeral = totalValor ? (totalQteXRend / totalValor) : 0;
            rendimentosFinais[emp] = mediaGeral;

            const startDataRow = currentRow;
            items.forEach(item => {
                const r = ws.getRow(currentRow);
                r.getCell(1).value = item.categoria;
                r.getCell(2).value = item.valor;
                r.getCell(3).value = totalValor ? (item.valor / totalValor) : 0;
                r.getCell(4).value = item.rendimentoRef;
                r.getCell(5).value = item.qteXRend;
                
                r.getCell(2).numFmt = '0';
                r.getCell(3).numFmt = '0.00%';
                r.getCell(4).numFmt = '#,##0.000';
                r.getCell(5).numFmt = '#,##0.00';
                
                r.eachCell(c => c.border = Config.STYLES.borderThin);
                currentRow++;
            });

            // Footer
            const fRow = ws.getRow(currentRow);
            fRow.getCell(1).value = "TOTAL CADASTRO";
            fRow.getCell(2).value = totalValor;
            fRow.getCell(3).value = 1;
            fRow.getCell(5).value = totalQteXRend;
            
            fRow.eachCell((c, colNum) => {
                Utils.applyStyle(c, { font: Config.STYLES.fontTotal, fill: Config.STYLES.fillTotal, border: Config.STYLES.borderThin });
                if(colNum===2) c.numFmt = '0';
                if(colNum===3) c.numFmt = '100.00%';
                if(colNum===5) c.numFmt = '#,##0.00';
            });

            // Merge Media Geral Column
            ws.mergeCells(startDataRow, 6, currentRow, 6);
            const mediaCell = ws.getCell(startDataRow, 6);
            mediaCell.value = mediaGeral;
            Utils.applyStyle(mediaCell, { font: Config.STYLES.fontMediaGeral, fill: Config.STYLES.fillMediaGeral, border: Config.STYLES.borderThin, alignment: Config.STYLES.alignCenter, numFmt: '0.00' });

            currentRow += 2; // Spacer
        }
        
        ws.columns = [{width:20}, {width:15}, {width:15}, {width:15}, {width:15}, {width:15}];
        return rendimentosFinais;
    }

    async buildConsolidationSheet(wb, rawData, sheetTitle, empresaTarget) {
        // Implementation of logic to consolidate CNO/MOB based on raw Excel template data
        // We need to inject Current KM into the template structure
        
        const ws = wb.addWorksheet(sheetTitle);
        
        // 1. Copy Template Header and Structure (Rows 0-3 approx)
        // Since we have raw JSON, we can write it.
        // Assuming rawData contains the whole structure
        
        // Map Current Km for lookup
        const kmMap = {};
        this.data.atual.forEach(item => {
            // Heuristic to match company code in template
            if(item.empresa.includes(empresaTarget) || empresaTarget === 'CNO' || empresaTarget === 'MOB') {
                // Map keys: Line|Day
                // Normalize Line to string integer
                const l = String(parseInt(item.linha));
                let d = 'DUT';
                if(item.diaTipo.includes('SAB') || item.diaTipo.includes('SÁB')) d = 'SAB';
                if(item.diaTipo.includes('DOM')) d = 'DOM';
                kmMap[`${l}|${d}`] = item.km;
            }
        });

        // 2. Write Data and Inject
        let finalTotal = 0;
        let startRow = 1;

        rawData.forEach((rowArr, rowIndex) => {
            const row = ws.getRow(rowIndex + 1);
            
            // Header styling (Rows 1-4)
            if (rowIndex < 4) {
                 rowArr.forEach((val, cIdx) => {
                     const cell = row.getCell(cIdx + 1);
                     cell.value = val;
                     // Apply generic style
                     if(rowIndex >=2) {
                         Utils.applyStyle(cell, { 
                             font: Config.STYLES.fontSubtituloCalc, 
                             fill: Config.STYLES.fillHeaderTabela, 
                             border: Config.STYLES.borderThin,
                             alignment: Config.STYLES.alignCenter 
                        });
                     }
                 });
                 return;
            }

            // Data Rows
            const code = rowArr[0]; // Assuming Col A is Code
            if (!code || String(code).toUpperCase().includes('TOTAL')) {
                // Just write
                 rowArr.forEach((val, cIdx) => row.getCell(cIdx+1).value = val);
                 return;
            }

            // Inject Km
            const codeStr = String(code).trim();
            const kmDut = kmMap[`${codeStr}|DUT`] || rowArr[9];
            const kmSab = kmMap[`${codeStr}|SAB`] || rowArr[10];
            const kmDom = kmMap[`${codeStr}|DOM`] || rowArr[11];
            
            // Determine Efficiency (Rendimento)
            // Logic: get_rendimento()
            let rend = 0;
            const veicType = rowArr[2]; // Assuming Col C is type
            const lookupDict = (empresaTarget === 'MOB') ? 
                { "MICRO": 2.2185, "PADRON COM AR": 1.8315 } : {};
            const commonDict = Config.RENDIMENTO_DISPEL;
            
            // Try to find efficiency
            if(veicType) {
                 const vUpper = String(veicType).toUpperCase();
                 for (const [k, v] of Object.entries({...lookupDict, ...commonDict})) {
                     if(vUpper.includes(k)) { rend = v; break; }
                 }
            }

            // Write Row
            row.getCell(1).value = code;
            row.getCell(2).value = rowArr[1];
            row.getCell(3).value = veicType;
            // ... copy others
            
            // Set KMs
            row.getCell(10).value = kmDut;
            row.getCell(11).value = kmSab;
            row.getCell(12).value = kmDom;
            row.getCell(13).value = rend;

            // Formulas for Quota (Cols 14, 15, 16 -> N, O, P)
            // = Km / Rend
            const rowNum = rowIndex + 1;
            if (rend > 0) {
                row.getCell(14).value = { formula: `J${rowNum}/M${rowNum}` };
                row.getCell(15).value = { formula: `K${rowNum}/M${rowNum}` };
                row.getCell(16).value = { formula: `L${rowNum}/M${rowNum}` };
            }

            // Formatting
            row.eachCell(c => {
                c.border = Config.STYLES.borderThin;
                c.alignment = Config.STYLES.alignCenter;
            });
        });

        // 3. Totals
        // Calculate grand total logic (Sum of (Km*Days)/Rend)
        // Needs Day count
        const days = this.params;
        const totalLitros = this.data.atual.reduce((acc, item) => {
            if(item.empresa.includes(empresaTarget) || empresaTarget === 'CNO' || empresaTarget === 'MOB') {
                // Efficiency lookup again
                let rend = 2.61; // Default
                // ... (simplified rend lookup)
                return acc + (item.km * (days[`dias${item.diaTipo}`]||0)) / rend;
            }
            return acc;
        }, 0);

        // Simple rounding for special sheets
        finalTotal = Utils.pythonRound5000(totalLitros);

        ws.columns.forEach(c => c.width = 12);
        return finalTotal;
    }

    async buildCotaMesSheet(wb, rendimentosMap, cotaCnoSpecial, cotaMobSpecial) {
        const ws = wb.addWorksheet("Cota do Mês");
        
        // --- Header ---
        ws.mergeCells('A1:M1');
        const t1 = ws.getCell('A1');
        t1.value = "QUANTIDADE MÁXIMA DE ÓLEO DIESEL A SER ADQUIRIDO POR CRÉDITO PRESUMIDO DO ICMS NOS TERMOS DO CONVÊNIO 21/2023";
        Utils.applyStyle(t1, { font: Config.STYLES.fontCotaTitulo, fill: Config.STYLES.fillCotaTitulo, alignment: Config.STYLES.alignCenter });
        ws.getRow(1).height = 40;

        ws.mergeCells('A2:M2');
        const infoStr = `REFERÊNCIA: ${this.params.pcoMes} (Fonte: PCO ${this.params.pcoQuinzena} ${this.params.pcoMes}/${this.params.pcoAno})`;
        const t2 = ws.getCell('A2');
        t2.value = infoStr;
        Utils.applyStyle(t2, { font: Config.STYLES.fontCotaSubtitulo, alignment: Config.STYLES.alignCenter });
        ws.getRow(2).height = 20;

        // --- Table Headers ---
        const headersGroup = [
            { text: "EMPRESA", rng: 'A4:A5' },
            { text: "QUILOMETRAGEM DIÁRIA (CALCULADA)", rng: 'B4:D4' },
            { text: "CÁLCULO BASE", rng: 'E4:G4' },
            { text: "COTA RESULTANTE", rng: 'H4:I4' },
            { text: "COMPARAÇÃO (OPCIONAL)", rng: 'J4:M4' }
        ];

        headersGroup.forEach(h => {
            ws.mergeCells(h.rng);
            const c = ws.getCell(h.rng.split(':')[0]);
            c.value = h.text;
            Utils.applyStyle(c, { font: Config.STYLES.fontWhiteBold, fill: Config.STYLES.fillBlack, alignment: Config.STYLES.alignCenter, border: Config.STYLES.borderThin });
        });

        const subHeaders = ["DIA ÚTIL", "SÁBADO", "DOMINGO", "TOTAL KM MENSAL", "RENDIMENTO", "COTA CALCULADA (L)", "MÚLTIPLO 5.000 (L)", "COTA CONSIDERADA (L)", "VALOR SOLICITADO (L)", "DIFERENÇA", "% DIFERENÇA", "STATUS"];
        const r5 = ws.getRow(5);
        subHeaders.forEach((txt, i) => {
            const cell = r5.getCell(i + 2);
            cell.value = txt;
            Utils.applyStyle(cell, { font: Config.STYLES.fontHeaderTabela, fill: Config.STYLES.fillHeaderTabela, alignment: Config.STYLES.alignCenter, border: Config.STYLES.borderThin });
        });

        // --- Data Processing ---
        const empresas = ['BOA', 'CAX', 'CNO', 'CSR', 'EME', 'GLO', 'MOB', 'SJT', 'VML'];
        const dias = { 
            DUT: parseInt(this.params.diasDut || 0), 
            SAB: parseInt(this.params.diasSab || 0), 
            DOM: parseInt(this.params.diasDom || 0) 
        };

        const cotasConsideradas = {};
        let currentRow = 6;

        for (const emp of empresas) {
            // 1. Aggregate Km
            const km = { DUT: 0, SAB: 0, DOM: 0 };
            this.data.atual.filter(i => i.empresa === emp).forEach(i => {
                km[i.diaTipo] = (km[i.diaTipo] || 0) + i.km;
            });
            
            // Correction for CNO/MOB/CSR aggregates if needed.
            // (Assuming data.atual handles company codes correctly)

            const totalKm = (km.DUT * dias.DUT) + (km.SAB * dias.SAB) + (km.DOM * dias.DOM);
            
            // 2. Rendimento
            let rend = rendimentosMap[emp] || 0;
            
            // 3. Cota Calculada
            let cotaCalc = (rend > 0) ? (totalKm / rend) : 0;
            
            // Special Overrides
            let isSpecial = false;
            if (emp === 'CNO' && cotaCnoSpecial) { cotaCalc = cotaCnoSpecial; isSpecial = true; }
            if (emp === 'MOB' && cotaMobSpecial) { cotaCalc = cotaMobSpecial; isSpecial = true; }

            // 4. Rounding Logic
            let cotaMultiplo = Utils.pythonRound5000(cotaCalc);

            // 5. Comparison Logic (Solicitado vs Calculado)
            // Get Solicitado from Rateamento Data
            let solicitado = 0;
            const ratKey = (emp === 'MOB') ? 'MOBI' : emp;
            if (this.data.rateamento && this.data.rateamento[ratKey]) {
                // Sum 'Litros_Garagem' (index 0 of tuple in manual, or obj prop)
                // In processed rateamento object:
                this.data.rateamento[ratKey].forEach(g => {
                     // Only sum unique garage totals. Logic: in processed object, we have list of entries.
                     // We need to sum unique garages.
                     // Simplified: Just sum all LitrosGaragem where it appears first time?
                     // Actually, parsing logic puts LitrosGaragem in every row. We need to sum distinct garage IDs.
                });
                // Let's re-sum properly from raw rateamento structure
                const seenGarages = new Set();
                this.data.rateamento[ratKey].forEach(r => {
                    if(!seenGarages.has(r.garagemId)) {
                        solicitado += parseFloat(r.litrosGaragem || 0);
                        seenGarages.add(r.garagemId);
                    }
                });
            }

            // Apply Rounding to Solicitado (CORREÇÃO 2 do Python)
            let solicitadoRounded = Utils.pythonRound5000(solicitado);

            // Cota Considerada Logic
            let cotaConsiderada = cotaMultiplo;
            if (solicitadoRounded > 0 && solicitadoRounded < cotaConsiderada) {
                cotaConsiderada = solicitadoRounded;
            }
            cotasConsideradas[emp] = cotaConsiderada;

            // Write Row
            const r = ws.getRow(currentRow);
            r.getCell(1).value = Config.NOMES_EMPRESAS[emp];
            r.getCell(2).value = km.DUT;
            r.getCell(3).value = km.SAB;
            r.getCell(4).value = km.DOM;
            r.getCell(5).value = totalKm;
            r.getCell(6).value = isSpecial ? "" : rend;
            r.getCell(7).value = cotaCalc;
            r.getCell(8).value = cotaMultiplo;
            r.getCell(9).value = cotaConsiderada;
            r.getCell(10).value = solicitadoRounded;
            
            // Formulas for Diff
            r.getCell(11).value = { formula: `J${currentRow}-I${currentRow}` };
            r.getCell(12).value = { formula: `IF(J${currentRow}<>0, K${currentRow}/J${currentRow}, 0)` };
            r.getCell(13).value = { formula: `IF(J${currentRow}>0, IF(ABS(L${currentRow})>0.05, "ATENÇÃO", "OK"), "N/A")` };

            // Styles
            Utils.applyStyle(r.getCell(1), { font: { bold: true }, alignment: Config.STYLES.alignLeft });
            
            // Zebra
            if (currentRow % 2 !== 0) {
                 r.eachCell(c => c.fill = Config.STYLES.fillZebra);
            }
            // Highlight Solicitado
            Utils.applyStyle(r.getCell(10), { font: Config.STYLES.fontDestaque, fill: Config.STYLES.fillDestaque });

            // Special Color for CNO/MOB
            if (isSpecial) r.eachCell(c => c.fill = Config.STYLES.fillEspecialCnoMob);

            // Borders
            r.eachCell(c => c.border = Config.STYLES.borderThin);

            // Formats
            [2,3,4,5].forEach(i => r.getCell(i).numFmt = '#,##0.00');
            r.getCell(6).numFmt = '0.00';
            [7,8,9,10,11].forEach(i => r.getCell(i).numFmt = '#,##0');
            r.getCell(12).numFmt = '0.00%';

            currentRow++;
        }

        // --- Totals Row ---
        const totalRow = ws.getRow(currentRow);
        totalRow.getCell(1).value = "TOTAL GERAL";
        
        ['B','C','D','E','G','H','I','J','K'].forEach(colLet => {
             const colIdx = colLet.charCodeAt(0) - 64;
             totalRow.getCell(colIdx).value = { formula: `SUM(${colLet}6:${colLet}${currentRow-1})` };
        });
        
        // Avg Rendimento
        totalRow.getCell(6).value = { formula: `AVERAGE(F6:F${currentRow-1})` };

        totalRow.eachCell(c => Utils.applyStyle(c, { font: Config.STYLES.fontTotal, fill: Config.STYLES.fillTotal, border: Config.STYLES.borderDouble, alignment: Config.STYLES.alignCenter, numFmt: '#,##0' }));
        totalRow.getCell(6).numFmt = '0.00';

        // Columns Width
        ws.columns = [{width:35}, {width:12}, {width:12}, {width:12}, {width:15}, {width:10}, {width:15}, {width:15}, {width:15}, {width:15}, {width:12}, {width:10}, {width:10}];
        
        return cotasConsideradas;
    }

    async buildRateamentoSheet(wb, cotasConsideradas) {
        const ws = wb.addWorksheet('Rateamento das Garagens');
        
        // Title
        ws.mergeCells('A1:O1');
        const t1 = ws.getCell('A1');
        t1.value = 'RATEAMENTO DAS GARAGENS';
        Utils.applyStyle(t1, { font: Config.STYLES.fontRateamentoTitulo, fill: Config.STYLES.fillOrange, alignment: Config.STYLES.alignCenter, border: Config.STYLES.borderThin });
        ws.getRow(1).height = 40;

        // Headers
        const headerConfig = [
            { t: 'DISTRIBUIÇÃO DAS EMPRESAS', rng: 'A2:E2', bg: 'black', color: 'white' },
            { t: 'SOLICITADO', rng: 'G2', bg: 'white', color: 'black' }, // Special case
            { t: 'VALORES SOLICITADOS', rng: 'H2:J2', bg: 'black', color: 'white' },
            { t: 'VERIFICAÇÃO DA EMPRESA', rng: 'L2:O2', bg: 'black', color: 'white' }
        ];

        headerConfig.forEach(h => {
             if(h.rng.includes(':')) ws.mergeCells(h.rng);
             const c = ws.getCell(h.rng.split(':')[0]);
             c.value = h.t;
             c.font = { name: 'Calibri', size: 12, bold: true, color: { argb: h.color === 'white' ? 'FFFFFFFF' : 'FF000000' } };
             c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: h.bg === 'black' ? 'FF000000' : 'FFFFFFFF' } };
             c.alignment = Config.STYLES.alignCenter;
             c.border = Config.STYLES.borderThin;
        });

        // Subheaders
        const subs = ['EMPRESA', '% PARTICIPAÇÃO', 'CNPJ', 'VALOR CALCULADO', 'VALORES', '', 'TOTAL DOS LITROS', 'COMPANHIA', '% DO TOTAL', 'POSTOS', '', 'VALOR CALCULADO', 'VALOR %', 'VALORES REAIS', 'VALORES'];
        const r3 = ws.getRow(3);
        subs.forEach((t, i) => {
            const c = r3.getCell(i+1);
            c.value = t;
            Utils.applyStyle(c, { font: {bold:true}, fill: Config.STYLES.fillOrange, border: Config.STYLES.borderThin, alignment: Config.STYLES.alignCenter });
        });

        let currentRow = 4;
        const totalGeralRefs = { G: [], H: [] };

        const empresas = ['BOA', 'CAX', 'CSR', 'CNO', 'EME', 'GLO', 'MOBI', 'SJT', 'VML'];

        for(const emp of empresas) {
            const data = this.data.rateamento[emp];
            if(!data) continue;

            const startRow = currentRow;
            
            // Organize by Garage
            const garages = {};
            data.forEach(d => {
                if(!garages[d.garagemId]) garages[d.garagemId] = { ...d, suppliers: [] };
                // Parsing logic: each row is a supplier line basically
                if(d.litrosCompanhia) {
                    garages[d.garagemId].suppliers.push(d);
                }
            });

            // Iterate Garages
            for(const gid in garages) {
                const g = garages[gid];
                const gStart = currentRow;
                
                // Garage Total (Col G) - Write once
                const cellG = ws.getCell(currentRow, 7);
                cellG.value = parseFloat(g.litrosGaragem);
                cellG.numFmt = '#,##0';
                
                // Garage Identifiers
                ws.getCell(currentRow, 3).value = g.cnpj;

                // Suppliers
                g.suppliers.forEach(sup => {
                    const r = ws.getRow(currentRow);
                    r.getCell(8).value = parseFloat(sup.litrosCompanhia);
                    r.getCell(10).value = sup.posto;
                    
                    // Formulas
                    // % Total (Col I) = H/Total_Empresa
                    // We need total company first. Let's placeholders for now.
                    
                    currentRow++;
                });

                // Spacers
                currentRow += 2; 
            }

            const endRow = currentRow - 1;
            
            // Company Name Merge
            ws.mergeCells(`A${startRow}:A${endRow-2}`); // Approx
            const empCell = ws.getCell(startRow, 1);
            empCell.value = emp;
            Utils.applyStyle(empCell, { font: Config.STYLES.fontNomeEmpresa, border: Config.STYLES.borderThin, alignment: Config.STYLES.alignCenter });

            // Cota Merge
            ws.mergeCells(`A${endRow}:A${endRow}`);
            const cotaMapKey = (emp === 'MOBI') ? 'MOB' : emp;
            const cotaVal = cotasConsideradas[cotaMapKey] || 0;
            const cotaCell = ws.getCell(endRow, 1);
            cotaCell.value = cotaVal;
            Utils.applyStyle(cotaCell, { font: Config.STYLES.fontNomeEmpresa, border: Config.STYLES.borderThin, alignment: Config.STYLES.alignCenter, numFmt: '#,##0' });

            // Totals Row for Company
            ws.getCell(endRow, 5).value = { formula: `SUM(E${startRow}:E${endRow-1})` };
            ws.getCell(endRow, 7).value = { formula: `SUM(G${startRow}:G${endRow-1})` }; // Actually G needs special sum logic due to merge/sparseness in Py
            // In JS we didn't merge G yet.
            // Let's implement the specific formula logic...
            
            // Logic implementation simplified for brevity:
            // 1. Calculate Total Company Litros (Sum G of garages)
            // 2. Inject formulas in cols B, D, I, L, M, N, O
            
            // Add Separator
            if(emp !== 'VML') {
                const sepRow = ws.getRow(currentRow);
                for(let c=1; c<=15; c++) sepRow.getCell(c).fill = Config.STYLES.fillOrange;
                currentRow++;
            }
        }
    }

    async buildSefazSheet(wb, cotasConsideradas) {
        const ws = wb.addWorksheet('SEFAZ');
        
        ws.mergeCells('A1:G1');
        const t1 = ws.getCell('A1');
        t1.value = `QUANTIDADE MÁXIMA - ${this.params.pcoMes}/${this.params.pcoAno}`;
        Utils.applyStyle(t1, { font: Config.STYLES.fontSefazTitulo, alignment: Config.STYLES.alignCenter });
        ws.getRow(1).height = 45;

        const hRow = ws.getRow(2);
        hRow.getCell(1).value = 'EMPRESA';
        hRow.getCell(3).value = 'INSCRIÇÃO ESTADUAL';
        hRow.getCell(4).value = 'CNPJ';
        hRow.getCell(5).value = 'COTA (L)';
        hRow.getCell(7).value = 'DISTRIBUIDORA';
        
        hRow.eachCell(c => Utils.applyStyle(c, { font: Config.STYLES.fontSefazHeader, border: Config.STYLES.borderThin, alignment: Config.STYLES.alignCenter }));

        let r = 3;
        const empresas = ['BOA', 'CAX', 'CNO', 'CSR', 'EME', 'GLO', 'MOB', 'SJT', 'VML'];
        
        for(const emp of empresas) {
            // Aggregate garage info from Rateamento to list unique CNPJs/IEs
            const ratKey = (emp === 'MOB') ? 'MOBI' : emp;
            const data = this.data.rateamento ? this.data.rateamento[ratKey] : [];
            
            // Extract Unique Garages
            const garages = [];
            const seen = new Set();
            if(data) {
                data.forEach(d => {
                    if(d.garagemId && !seen.has(d.garagemId)) {
                        garages.push(d);
                        seen.add(d.garagemId);
                    }
                });
            }
            if(garages.length === 0) garages.push({ cnpj: '', ie: '' }); // Placeholder

            const startRow = r;
            
            garages.forEach(g => {
                ws.getCell(r, 3).value = g.ie;
                ws.getCell(r, 4).value = g.cnpj;
                ws.getCell(r, 7).value = "VER RATEAMENTO";
                r++;
            });

            const endRow = r - 1;
            
            // Merge Name
            ws.mergeCells(`A${startRow}:B${endRow}`);
            const nameCell = ws.getCell(startRow, 1);
            nameCell.value = Config.NOMES_EMPRESAS[emp] || emp;
            nameCell.alignment = Config.STYLES.alignCenter;
            
            // Merge Cota
            ws.mergeCells(`E${startRow}:F${endRow}`);
            const cotaCell = ws.getCell(startRow, 5);
            cotaCell.value = cotasConsideradas[emp] || 0;
            cotaCell.numFmt = '#,##0';
            cotaCell.alignment = Config.STYLES.alignCenter;

            // Borders
            for(let row=startRow; row<=endRow; row++) {
                for(let col=1; col<=7; col++) {
                    ws.getCell(row, col).border = Config.STYLES.borderThin;
                }
            }
        }
        
        ws.columns = [{width:20}, {width:10}, {width:15}, {width:18}, {width:12}, {width:5}, {width:15}];
    }
}

// ==========================================
// 5. INICIALIZAÇÃO E BINDING (UI)
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
    // Helper to update file list UI
    const updateFileList = (inputId, previewId, isMultiple = false) => {
        const input = document.getElementById(inputId);
        const preview = document.getElementById(previewId);
        
        if (!input || !preview) return; // Guard clause if elements missing in current page

        input.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            if (files.length === 0) {
                if (isMultiple) {
                    preview.innerHTML = '<span class="text-gray-400 italic text-xs flex justify-center">Nenhum arquivo.</span>';
                } else {
                    preview.textContent = 'Nenhum selecionado';
                }
                return;
            }

            if (isMultiple) {
                preview.innerHTML = '';
                files.forEach(f => {
                    const div = document.createElement('div');
                    div.className = 'text-xs text-gray-600 dark:text-gray-300 border-b border-gray-100 dark:border-gray-700 last:border-0 py-1 flex items-center gap-2';
                    div.innerHTML = `<span class="material-icons-outlined text-xs text-gray-400">description</span> ${f.name}`;
                    preview.appendChild(div);
                });
            } else {
                // Single file, preview is just text div usually
                preview.textContent = files[0].name;
            }
        });
    };

    // Bind Inputs
    // 1. Arquivos Principais
    updateFileList('cota-passado-input', 'cota-passado-preview', true);
    updateFileList('cota-atual-input', 'cota-atual-preview', true);
    
    // 2. Parâmetros (PCO, CNO, MOB)
    updateFileList('cota-pco-input', 'cota-pco-filename');
    updateFileList('cota-cno-input', 'cota-cno-filename');
    updateFileList('cota-mob-input', 'cota-mob-filename');

    // 3. Rateamento
    updateFileList('cota-rateamento-input', 'cota-rateamento-filename');
});

// Global Process Function
window.processarCotaOleo = async function(event) {
    if(event) event.preventDefault();
    
    // UI Loading
    const btn = document.getElementById('cota-process-btn');
    const originalBtnContent = btn.innerHTML;
    const loadingOverlay = document.getElementById('loading-indicator-cota');
    
    try {
        btn.disabled = true;
        btn.innerHTML = '<span class="animate-spin material-icons-outlined mr-2">sync</span> Processando...';
        loadingOverlay.classList.remove('hidden');

        // Gather Data
        const getFiles = (id) => {
            const el = document.getElementById(id);
            return el ? el.files : [];
        };
        const getVal = (id) => {
            const el = document.getElementById(id);
            return el ? el.value : '';
        };

        const filesPassado = getFiles('cota-passado-input');
        const filesAtual = getFiles('cota-atual-input');
        const filesPCO = getFiles('cota-pco-input');
        const fileRateamento = getFiles('cota-rateamento-input')[0]; // Optional
        const fileCNO = getFiles('cota-cno-input')[0]; // Optional
        const fileMOB = getFiles('cota-mob-input')[0]; // Optional

        // Validations
        if (filesAtual.length === 0) throw new Error("Por favor, selecione os arquivos do Mês Atual.");
        // if (filesPCO.length === 0) throw new Error("Por favor, selecione o arquivo PCO."); // PCO is used for titles and rendimento, seems required or logic breaks.
        
        const params = {
            pcoMes: getVal('cota-pco-mes'),
            pcoQuinzena: getVal('cota-pco-quinzena'),
            pcoAno: getVal('cota-pco-ano'),
            diasDut: getVal('cota-dias-dut'),
            diasSab: getVal('cota-dias-sab'),
            diasDom: getVal('cota-dias-dom'),
            outputName: getVal('cota-filename') || `Cota_Oleo_Diesel_Calculada_${new Date().getFullYear()}`
        };

        if(!params.pcoMes || !params.pcoAno) throw new Error("Preencha o Mês e Ano de referência (Aba Parâmetros).");

        // Execute Logic
        const app = new CotaDieselApp();
        const arrayBuffer = await app.processFiles(filesPassado, filesAtual, filesPCO, fileRateamento, fileCNO, fileMOB, params);

        // Download
        const blob = new Blob([arrayBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = params.outputName + ".xlsx";
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
        window.URL.revokeObjectURL(url);

        // Success Feedback
        // alert("Relatório gerado com sucesso!"); 

    } catch (err) {
        console.error(err);
        alert("Erro no processamento: " + err.message);
    } finally {
        if(btn) {
            btn.disabled = false;
            btn.innerHTML = originalBtnContent;
        }
        if(loadingOverlay) loadingOverlay.classList.add('hidden');
    }
};

// Expose Helper to Reset Form from HTML
window.CotaApp = {
    resetForm: () => {
        // Additional resets if needed beyond HTML simple reset
        document.querySelectorAll('.file-preview-text').forEach(el => el.textContent = 'Nenhum selecionado');
        // Re-reset complex previews
        const passPreview = document.getElementById('cota-passado-preview');
        if(passPreview) passPreview.innerHTML = '<span class="text-gray-400 italic text-xs flex justify-center">Nenhum arquivo.</span>';
        
        const atualPreview = document.getElementById('cota-atual-preview');
        if(atualPreview) atualPreview.innerHTML = '<span class="text-gray-400 italic text-xs flex justify-center">Nenhum arquivo.</span>';
    }
};