class Spreadsheet {
    constructor() {
        this.rows = 100;
        this.cols = 26;
        this.data = {};
        this.selectedCell = null;
        this.init();
    }

    init() {
        this.createTable();
        this.setupEventListeners();
    }

    createTable() {
        const headerRow = document.getElementById('headerRow');
        const tbody = document.getElementById('spreadsheetBody');

        // ヘッダー行の作成
        const cornerTh = document.createElement('th');
        cornerTh.className = 'row-header';
        headerRow.appendChild(cornerTh);

        for (let col = 0; col < this.cols; col++) {
            const th = document.createElement('th');
            th.textContent = this.getColumnName(col);
            headerRow.appendChild(th);
        }

        // データ行の作成
        for (let row = 0; row < this.rows; row++) {
            const tr = document.createElement('tr');
            
            // 行番号
            const rowHeader = document.createElement('th');
            rowHeader.className = 'row-header';
            rowHeader.textContent = row + 1;
            tr.appendChild(rowHeader);

            // セルの作成
            for (let col = 0; col < this.cols; col++) {
                const td = document.createElement('td');
                const input = document.createElement('input');
                input.type = 'text';
                input.dataset.row = row;
                input.dataset.col = col;
                input.id = `cell-${row}-${col}`;
                td.appendChild(input);
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }
    }

    getColumnName(col) {
        let name = '';
        while (col >= 0) {
            name = String.fromCharCode(65 + (col % 26)) + name;
            col = Math.floor(col / 26) - 1;
        }
        return name;
    }

    getCellId(row, col) {
        return `${this.getColumnName(col)}${row + 1}`;
    }

    parseCellId(cellId) {
        const match = cellId.match(/^([A-Z]+)(\d+)$/);
        if (!match) return null;

        const colName = match[1];
        const rowNum = parseInt(match[2]) - 1;

        let col = 0;
        for (let i = 0; i < colName.length; i++) {
            col = col * 26 + (colName.charCodeAt(i) - 64);
        }
        col--;

        return { row: rowNum, col: col };
    }

    setupEventListeners() {
        const tbody = document.getElementById('spreadsheetBody');
        const formulaInput = document.getElementById('formulaInput');

        // セルのクリックイベント
        tbody.addEventListener('click', (e) => {
            if (e.target.tagName === 'INPUT') {
                this.selectCell(e.target);
            }
        });

        // セルの値変更イベント
        tbody.addEventListener('input', (e) => {
            if (e.target.tagName === 'INPUT') {
                const row = parseInt(e.target.dataset.row);
                const col = parseInt(e.target.dataset.col);
                const value = e.target.value;
                this.setCellValue(row, col, value);
                this.updateFormulaBar();
            }
        });

        // キーボードイベント（Enter、カーソルキー）
        tbody.addEventListener('keydown', (e) => {
            if (e.target.tagName === 'INPUT') {
                const row = parseInt(e.target.dataset.row);
                const col = parseInt(e.target.dataset.col);
                let nextCell = null;

                switch(e.key) {
                    case 'Enter':
                        e.preventDefault();
                        this.evaluateCell(e.target);
                        const nextRow = Math.min(row + 1, this.rows - 1);
                        nextCell = document.getElementById(`cell-${nextRow}-${col}`);
                        break;
                    
                    case 'ArrowUp':
                        e.preventDefault();
                        if (row > 0) {
                            nextCell = document.getElementById(`cell-${row - 1}-${col}`);
                        }
                        break;
                    
                    case 'ArrowDown':
                        e.preventDefault();
                        if (row < this.rows - 1) {
                            nextCell = document.getElementById(`cell-${row + 1}-${col}`);
                        }
                        break;
                    
                    case 'ArrowLeft':
                        if (e.target.selectionStart === 0 && e.target.selectionEnd === 0) {
                            e.preventDefault();
                            if (col > 0) {
                                nextCell = document.getElementById(`cell-${row}-${col - 1}`);
                            }
                        }
                        break;
                    
                    case 'ArrowRight':
                        if (e.target.selectionStart === e.target.value.length) {
                            e.preventDefault();
                            if (col < this.cols - 1) {
                                nextCell = document.getElementById(`cell-${row}-${col + 1}`);
                            }
                        }
                        break;
                    
                    case 'Tab':
                        e.preventDefault();
                        if (e.shiftKey) {
                            // Shift+Tab で左へ
                            if (col > 0) {
                                nextCell = document.getElementById(`cell-${row}-${col - 1}`);
                            } else if (row > 0) {
                                nextCell = document.getElementById(`cell-${row - 1}-${this.cols - 1}`);
                            }
                        } else {
                            // Tab で右へ
                            if (col < this.cols - 1) {
                                nextCell = document.getElementById(`cell-${row}-${col + 1}`);
                            } else if (row < this.rows - 1) {
                                nextCell = document.getElementById(`cell-${row + 1}-0`);
                            }
                        }
                        break;
                }

                if (nextCell) {
                    nextCell.focus();
                    this.selectCell(nextCell);
                }
            }
        });

        // フォーミュラバーの入力
        formulaInput.addEventListener('input', (e) => {
            if (this.selectedCell) {
                this.selectedCell.value = e.target.value;
                const row = parseInt(this.selectedCell.dataset.row);
                const col = parseInt(this.selectedCell.dataset.col);
                this.setCellValue(row, col, e.target.value);
            }
        });

        formulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && this.selectedCell) {
                e.preventDefault();
                this.evaluateCell(this.selectedCell);
            }
        });

        // ボタンのイベント
        document.getElementById('saveBtn').addEventListener('click', () => this.saveData());
        document.getElementById('loadBtn').addEventListener('click', () => this.loadData());
        document.getElementById('clearBtn').addEventListener('click', () => this.clearData());
        document.getElementById('fileInput').addEventListener('change', (e) => this.handleFileSelect(e));
    }

    selectCell(input) {
        // 前の選択を解除
        if (this.selectedCell) {
            this.selectedCell.parentElement.classList.remove('selected');
        }

        this.selectedCell = input;
        input.parentElement.classList.add('selected');
        this.updateFormulaBar();
    }

    updateFormulaBar() {
        const formulaInput = document.getElementById('formulaInput');
        if (this.selectedCell) {
            const row = parseInt(this.selectedCell.dataset.row);
            const col = parseInt(this.selectedCell.dataset.col);
            const cellData = this.data[`${row}-${col}`];
            formulaInput.value = cellData ? cellData.formula || cellData.value : '';
            formulaInput.placeholder = this.getCellId(row, col);
        }
    }

    setCellValue(row, col, value) {
        const key = `${row}-${col}`;
        if (value.trim() === '') {
            delete this.data[key];
        } else {
            this.data[key] = {
                value: value,
                formula: value.startsWith('=') ? value : null
            };
        }
    }

    evaluateCell(input) {
        const row = parseInt(input.dataset.row);
        const col = parseInt(input.dataset.col);
        const cellData = this.data[`${row}-${col}`];

        if (cellData && cellData.formula) {
            try {
                const result = this.evaluateFormula(cellData.formula);
                input.value = result;
                input.parentElement.classList.add('formula');
                input.parentElement.classList.remove('error');
            } catch (error) {
                input.value = '#ERROR!';
                input.parentElement.classList.add('error');
                input.parentElement.classList.remove('formula');
            }
        } else {
            input.parentElement.classList.remove('formula', 'error');
        }
    }

    evaluateFormula(formula) {
        // 基本的な数式の評価
        let expression = formula.substring(1); // '='を除去

        // SUM関数の処理
        expression = expression.replace(/SUM\(([A-Z]+\d+):([A-Z]+\d+)\)/gi, (match, start, end) => {
            const startCell = this.parseCellId(start);
            const endCell = this.parseCellId(end);
            
            if (!startCell || !endCell) return '0';

            let sum = 0;
            for (let r = startCell.row; r <= endCell.row; r++) {
                for (let c = startCell.col; c <= endCell.col; c++) {
                    const cellData = this.data[`${r}-${c}`];
                    if (cellData && !cellData.formula) {
                        const value = parseFloat(cellData.value);
                        if (!isNaN(value)) sum += value;
                    }
                }
            }
            return sum;
        });

        // セル参照の処理
        expression = expression.replace(/([A-Z]+\d+)/g, (match, cellId) => {
            const cell = this.parseCellId(cellId);
            if (!cell) return '0';
            
            const cellData = this.data[`${cell.row}-${cell.col}`];
            if (cellData && !cellData.formula) {
                const value = parseFloat(cellData.value);
                return isNaN(value) ? '0' : value;
            }
            return '0';
        });

        // 安全な評価（基本的な演算のみ）
        try {
            // 基本的な数学演算のみを許可
            if (!/^[\d\s+\-*/().,]+$/.test(expression)) {
                throw new Error('Invalid expression');
            }
            return eval(expression);
        } catch (error) {
            throw error;
        }
    }

    saveData() {
        const dataToSave = {
            version: '1.0',
            data: this.data
        };
        const json = JSON.stringify(dataToSave, null, 2);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'spreadsheet_data.json';
        a.click();
        URL.revokeObjectURL(url);
    }

    loadData() {
        document.getElementById('fileInput').click();
    }

    handleFileSelect(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const loadedData = JSON.parse(event.target.result);
                this.data = loadedData.data || {};
                this.updateAllCells();
            } catch (error) {
                alert('ファイルの読み込みに失敗しました。');
            }
        };
        reader.readAsText(file);
    }

    clearData() {
        if (confirm('すべてのデータをクリアしますか？')) {
            this.data = {};
            this.updateAllCells();
        }
    }

    updateAllCells() {
        for (let row = 0; row < this.rows; row++) {
            for (let col = 0; col < this.cols; col++) {
                const input = document.getElementById(`cell-${row}-${col}`);
                const cellData = this.data[`${row}-${col}`];
                
                if (cellData) {
                    input.value = cellData.value;
                    if (cellData.formula) {
                        this.evaluateCell(input);
                    }
                } else {
                    input.value = '';
                    input.parentElement.classList.remove('formula', 'error');
                }
            }
        }
    }
}

// アプリケーションの初期化
document.addEventListener('DOMContentLoaded', () => {
    new Spreadsheet();
});