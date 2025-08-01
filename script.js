class Spreadsheet {
    constructor() {
        this.rows = 100;
        this.cols = 26;
        this.data = {};
        this.selectedCell = null;
        this.selectionStart = null;
        this.selectionEnd = null;
        this.isSelecting = false;
        this.rangeStartCell = null; // 範囲選択の起点を記録
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
                td.dataset.row = row;
                td.dataset.col = col;
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
                this.clearRangeSelection();
                this.rangeStartCell = null;
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
                        if (e.shiftKey) {
                            // Shift+上で範囲選択
                            if (!this.rangeStartCell) {
                                this.rangeStartCell = { row, col };
                            }
                            if (row > 0) {
                                this.selectRange(this.rangeStartCell.row, this.rangeStartCell.col, row - 1, col);
                                nextCell = document.getElementById(`cell-${row - 1}-${col}`);
                            }
                        } else {
                            if (row > 0) {
                                nextCell = document.getElementById(`cell-${row - 1}-${col}`);
                            }
                            this.clearRangeSelection();
                            this.rangeStartCell = null;
                        }
                        break;
                    
                    case 'ArrowDown':
                        e.preventDefault();
                        if (e.shiftKey) {
                            // Shift+下で範囲選択
                            if (!this.rangeStartCell) {
                                this.rangeStartCell = { row, col };
                            }
                            if (row < this.rows - 1) {
                                this.selectRange(this.rangeStartCell.row, this.rangeStartCell.col, row + 1, col);
                                nextCell = document.getElementById(`cell-${row + 1}-${col}`);
                            }
                        } else {
                            if (row < this.rows - 1) {
                                nextCell = document.getElementById(`cell-${row + 1}-${col}`);
                            }
                            this.clearRangeSelection();
                            this.rangeStartCell = null;
                        }
                        break;
                    
                    case 'ArrowLeft':
                        if (e.target.selectionStart === 0 && e.target.selectionEnd === 0) {
                            e.preventDefault();
                            if (e.shiftKey) {
                                // Shift+左で範囲選択
                                if (!this.rangeStartCell) {
                                    this.rangeStartCell = { row, col };
                                }
                                if (col > 0) {
                                    this.selectRange(this.rangeStartCell.row, this.rangeStartCell.col, row, col - 1);
                                    nextCell = document.getElementById(`cell-${row}-${col - 1}`);
                                }
                            } else {
                                if (col > 0) {
                                    nextCell = document.getElementById(`cell-${row}-${col - 1}`);
                                }
                                this.clearRangeSelection();
                                this.rangeStartCell = null;
                            }
                        }
                        break;
                    
                    case 'ArrowRight':
                        if (e.target.selectionStart === e.target.value.length) {
                            e.preventDefault();
                            if (e.shiftKey) {
                                // Shift+右で範囲選択
                                if (!this.rangeStartCell) {
                                    this.rangeStartCell = { row, col };
                                }
                                if (col < this.cols - 1) {
                                    this.selectRange(this.rangeStartCell.row, this.rangeStartCell.col, row, col + 1);
                                    nextCell = document.getElementById(`cell-${row}-${col + 1}`);
                                }
                            } else {
                                if (col < this.cols - 1) {
                                    nextCell = document.getElementById(`cell-${row}-${col + 1}`);
                                }
                                this.clearRangeSelection();
                                this.rangeStartCell = null;
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
                        this.clearRangeSelection();
                        this.rangeStartCell = null;
                        break;
                }

                if (nextCell) {
                    nextCell.focus();
                    this.selectCell(nextCell);
                }
            }
        });

        // Ctrl+C でコピー、Ctrl+V で貼り付け
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.key === 'c') {
                if (this.selectionStart && this.selectionEnd) {
                    e.preventDefault();
                    this.copyRangeToClipboard();
                } else if (this.selectedCell) {
                    // 単一セルの場合は通常のコピーを許可
                    return;
                }
            } else if (e.ctrlKey && e.key === 'v') {
                if (this.selectedCell) {
                    e.preventDefault();
                    this.pasteFromClipboard();
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
        document.getElementById('exportBtn').addEventListener('click', () => this.showExportModal());
        document.getElementById('fileInput').addEventListener('change', (e) => this.handleFileSelect(e));

        // モーダル関連のイベント
        const modal = document.getElementById('exportModal');
        const closeBtn = document.querySelector('.close');
        closeBtn.addEventListener('click', () => this.hideExportModal());
        window.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.hideExportModal();
            }
        });

        document.getElementById('exportHTML').addEventListener('click', () => this.exportAsHTML());
        document.getElementById('exportMarkdown').addEventListener('click', () => this.exportAsMarkdown());
        document.getElementById('copyToClipboard').addEventListener('click', () => this.copyToClipboard());
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

    selectRange(startRow, startCol, endRow, endCol) {
        this.clearRangeSelection();
        
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);
        
        this.selectionStart = { row: minRow, col: minCol };
        this.selectionEnd = { row: maxRow, col: maxCol };
        
        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const cell = document.getElementById(`cell-${row}-${col}`);
                if (cell) {
                    cell.parentElement.classList.add('range-selected');
                }
            }
        }
    }

    clearRangeSelection() {
        document.querySelectorAll('.range-selected').forEach(td => {
            td.classList.remove('range-selected');
        });
        this.selectionStart = null;
        this.selectionEnd = null;
    }

    showExportModal() {
        const modal = document.getElementById('exportModal');
        const rangeSpan = document.getElementById('selectedRange');
        
        if (this.selectionStart && this.selectionEnd) {
            const startCell = this.getCellId(this.selectionStart.row, this.selectionStart.col);
            const endCell = this.getCellId(this.selectionEnd.row, this.selectionEnd.col);
            rangeSpan.textContent = `${startCell}:${endCell}`;
        } else if (this.selectedCell) {
            const row = parseInt(this.selectedCell.dataset.row);
            const col = parseInt(this.selectedCell.dataset.col);
            rangeSpan.textContent = this.getCellId(row, col);
        } else {
            rangeSpan.textContent = 'なし';
        }
        
        modal.style.display = 'block';
        document.getElementById('exportOutput').value = '';
    }

    hideExportModal() {
        document.getElementById('exportModal').style.display = 'none';
    }

    getSelectedData() {
        const data = [];
        let startRow, endRow, startCol, endCol;
        
        if (this.selectionStart && this.selectionEnd) {
            startRow = this.selectionStart.row;
            endRow = this.selectionEnd.row;
            startCol = this.selectionStart.col;
            endCol = this.selectionEnd.col;
        } else if (this.selectedCell) {
            startRow = endRow = parseInt(this.selectedCell.dataset.row);
            startCol = endCol = parseInt(this.selectedCell.dataset.col);
        } else {
            return data;
        }
        
        for (let row = startRow; row <= endRow; row++) {
            const rowData = [];
            for (let col = startCol; col <= endCol; col++) {
                const cellData = this.data[`${row}-${col}`];
                const value = cellData ? cellData.value : '';
                rowData.push(value);
            }
            data.push(rowData);
        }
        
        return data;
    }

    exportAsHTML() {
        const data = this.getSelectedData();
        if (data.length === 0) {
            document.getElementById('exportOutput').value = 'データが選択されていません';
            return;
        }
        
        let html = '<table>\n';
        data.forEach(row => {
            html += '  <tr>\n';
            row.forEach(cell => {
                html += `    <td>${this.escapeHtml(cell)}</td>\n`;
            });
            html += '  </tr>\n';
        });
        html += '</table>';
        
        document.getElementById('exportOutput').value = html;
    }

    exportAsMarkdown() {
        const data = this.getSelectedData();
        if (data.length === 0) {
            document.getElementById('exportOutput').value = 'データが選択されていません';
            return;
        }
        
        let markdown = '';
        
        // ヘッダー行
        markdown += '| ' + data[0].map(cell => this.escapeMarkdown(cell)).join(' | ') + ' |\n';
        
        // 区切り行
        markdown += '|' + data[0].map(() => ' --- ').join('|') + '|\n';
        
        // データ行
        for (let i = 1; i < data.length; i++) {
            markdown += '| ' + data[i].map(cell => this.escapeMarkdown(cell)).join(' | ') + ' |\n';
        }
        
        // 1行だけの場合は区切り行を追加
        if (data.length === 1) {
            markdown = '| ' + data[0].map(cell => this.escapeMarkdown(cell)).join(' | ') + ' |\n';
            markdown += '|' + data[0].map(() => ' --- ').join('|') + '|\n';
        }
        
        document.getElementById('exportOutput').value = markdown;
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    escapeMarkdown(text) {
        return text.replace(/[|\\]/g, '\\$&');
    }

    copyRangeToClipboard() {
        const data = this.getSelectedData();
        if (data.length === 0) return;
        
        // タブ区切り形式に変換（Excelで貼り付け可能）
        const tsvText = data.map(row => row.join('\t')).join('\n');
        
        // クリップボードにコピー
        if (navigator.clipboard) {
            navigator.clipboard.writeText(tsvText).then(() => {
                // 視覚的フィードバック
                this.showCopyFeedback();
            }).catch(err => {
                console.error('クリップボードへのコピーに失敗しました:', err);
            });
        } else {
            // フォールバック: 古いブラウザ用
            const textarea = document.createElement('textarea');
            textarea.value = tsvText;
            textarea.style.position = 'fixed';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
            this.showCopyFeedback();
        }
    }

    showCopyFeedback() {
        // コピー完了の視覚的フィードバック
        const feedbackDiv = document.createElement('div');
        feedbackDiv.textContent = 'コピーしました';
        feedbackDiv.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background-color: #333;
            color: white;
            padding: 10px 20px;
            border-radius: 4px;
            z-index: 1000;
            font-size: 14px;
        `;
        document.body.appendChild(feedbackDiv);
        
        setTimeout(() => {
            document.body.removeChild(feedbackDiv);
        }, 1000);
    }

    pasteFromClipboard() {
        if (!this.selectedCell) return;
        
        const startRow = parseInt(this.selectedCell.dataset.row);
        const startCol = parseInt(this.selectedCell.dataset.col);
        
        if (navigator.clipboard) {
            navigator.clipboard.readText().then(text => {
                this.pasteData(text, startRow, startCol);
            }).catch(err => {
                console.error('クリップボードからの読み取りに失敗しました:', err);
            });
        } else {
            // フォールバック: プロンプトを使用
            const text = prompt('貼り付けるデータを入力してください:');
            if (text) {
                this.pasteData(text, startRow, startCol);
            }
        }
    }

    pasteData(text, startRow, startCol) {
        // タブ区切りデータを解析
        const lines = text.trim().split('\n');
        const data = lines.map(line => line.split('\t'));
        
        // 貼り付け
        for (let i = 0; i < data.length; i++) {
            const row = startRow + i;
            if (row >= this.rows) break; // 範囲外は無視
            
            for (let j = 0; j < data[i].length; j++) {
                const col = startCol + j;
                if (col >= this.cols) break; // 範囲外は無視
                
                const cellId = `cell-${row}-${col}`;
                const input = document.getElementById(cellId);
                if (input) {
                    input.value = data[i][j];
                    this.setCellValue(row, col, data[i][j]);
                    
                    // 数式の場合は評価
                    if (data[i][j].startsWith('=')) {
                        this.evaluateCell(input);
                    }
                }
            }
        }
        
        // 貼り付け完了の視覚的フィードバック
        this.showPasteFeedback();
    }

    showPasteFeedback() {
        const feedbackDiv = document.createElement('div');
        feedbackDiv.textContent = '貼り付けました';
        feedbackDiv.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background-color: #333;
            color: white;
            padding: 10px 20px;
            border-radius: 4px;
            z-index: 1000;
            font-size: 14px;
        `;
        document.body.appendChild(feedbackDiv);
        
        setTimeout(() => {
            document.body.removeChild(feedbackDiv);
        }, 1000);
    }

    copyToClipboard() {
        const output = document.getElementById('exportOutput');
        output.select();
        document.execCommand('copy');
        
        const button = document.getElementById('copyToClipboard');
        const originalText = button.textContent;
        button.textContent = 'コピーしました！';
        setTimeout(() => {
            button.textContent = originalText;
        }, 2000);
    }
}

// アプリケーションの初期化
document.addEventListener('DOMContentLoaded', () => {
    new Spreadsheet();
});