// Генерация главного отчета.
// Главный отчет содержит данные за все месяцы в компактном виде.

// В пространстве имен reports.
if (typeof reports === 'undefined') {
    reports = {};
}

/**
 * @param {Object} options Настройки.
 *
 * @option options {String} title Наименование заголовка.
 *
 * @constructor
 * @author Vyatcheslav Suharnikov <sv@graffity.biz>
 */
reports.General = function (options) {
    if (!options) {
        options = {};
    }

    this._title = options.title || 'Отчет'; // Наименование для листа с отчетом.
    this._sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this._title);
};

reports.General.prototype = {
    ensureFirst: function () {
        reports.Utils.ensureSheetIndex(this._sheet, 0);
    },

    hideDays: function () {
        var sheet = this._sheet;

        // getMaxRows
    },

    hideMinutes: function () {
        var sheet = this._sheet;

        sheet.hideColumn(sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1));
    },

    ensureReport: function () {
        if (!this._sheet) {
            this._sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(this._title, 0);
            this._prepareSheet(); // delete this
        }

        this.hideDays();
        this.ensureFirst();
    },

    _prepareSheet: function () {
        this._createHeader(1, 1);
    },

    _createHeader: function (startRow, startColumn) {
        var sheet = this._sheet,
            columnWidths = [
                75,     // Дата.
                120,    // Проект.
                460,    // Задача.
                65,     // Начало.
                65,     // Конец.
                65,     // Минут.
                580     // Комментарий.
            ],
            columnCount = columnWidths.length;

        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            sheet.setColumnWidth(startColumn + columnIndex, columnWidths[columnIndex]);
        }

        sheet.getRange(startRow, startColumn, 1, columnCount)
            .setValues([
                ['Дата', 'Проект', 'Задача', 'Начало', 'Конец', 'Минут', 'Комментарий']
            ])
            .setFontWeight('bold')
            .setHorizontalAlignment('center')
            .setBackground('#969696')
            .setFontColor('#fff')
            .setFontSize(12);

        sheet.setFrozenRows(startRow);

        // Задаем форматирование ячеек.
        var lastRow = sheet.getMaxRows() - startRow;

        // Дата.
        sheet.getRange(startRow + 1, startColumn, lastRow, 1)
            .setNumberFormat('dd.MM.YYYY')
            .setHorizontalAlignment('center');

        // Проект.
        sheet.getRange(startRow + 1, startColumn + 1, lastRow, 1)
            .setHorizontalAlignment('center');

        // Начало, Конец.
        sheet.getRange(startRow + 1, startColumn + 3, lastRow, 2)
            .setNumberFormat('00:00')
            .setHorizontalAlignment('center');

        // Формулы.
        sheet.getRange(startRow + 1, startColumn + 5, lastRow, 1)
            .setFormulaR1C1('=HOUR(R[0]C[-1]-R[0]C[-2])*60 + MINUTE(R[0]C[-1]-R[0]C[-2])');

        // Скрываем столбец с минутами.
        this.hideMinutes();

        // Удаляем лишние столбцы.
        if (sheet.getMaxColumns() - columnCount > 0) {
            sheet.deleteColumns(startColumn + columnCount, sheet.getMaxColumns() - columnCount);
        }
    }
};

reports.general = new reports.General();