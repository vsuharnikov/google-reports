// Генерация ежемесячного отчета.
// В пространстве имен reports.
if (typeof reports === 'undefined') {
    reports = {};
}

/**
 * @param {Date} date Дата, для которой создается отчет.
 *
 * @constructor
 * @author Vyatcheslav Suharnikov <sv@graffity.biz>
 */
reports.Monthly = function (date) {
    this._date = date;
    this._daysInMonth = reports.Utils.getDaysInMonth(date);
    this._reportName = reports.Utils.formatDateMMMMY(date); // Наименование для листа с отчетом.
    this._sheet = null;
};

reports.Monthly.prototype = {
    /**
     * Создание пустого отчета для месяца, указанного в дате.
     *
     * Отчет создается в новой вкладке на первом месте.
     * Если отчет уже создан - появляется предупреждение и осуществляется переход к отчету.
     */
    makeEmptyReport: function () {
        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(), // Текущий документ.
            sheet = spreadsheet.getSheetByName(this._reportName);

        if (sheet) {
            // Уже готов.
            sheet.activate();
            SpreadsheetApp.getUi().alert('Отчет "' + this._reportName + '" уже создан.');
            return;
        }

        this._sheet = spreadsheet.insertSheet(this._reportName, 0);

        this._initialize();

        this._setupHeader(1, 1);
        this._setupDayRows(4, 1); // Первые 3 строки для заголовка.
        this._setupSummaryRow(4 + this._daysInMonth, 1);
        this._setupFooter(4 + this._daysInMonth + 1, 1);
        this._setupLegend(1, 18);

        // Формула для проверки.
        var summaryRow = 4 + this._daysInMonth;
        var summaryResultCell = this._sheet.getRange(summaryRow, 15).getA1Notation();
        var checkFormula = '=';
        var checkLegendFormula = '=';

        checkFormula += summaryResultCell + '-' + this._sheet.getRange(summaryRow, 11).getA1Notation();
        checkLegendFormula += '' + summaryResultCell + '-' + this._sheet.getRange(2 + this._daysInMonth, 18).getA1Notation();

        this._setupCheckArea(summaryRow, 19, checkFormula, checkLegendFormula);
    },

    /**
     * Подгоняет размеры листа под необходимые, а так же применяет общие стили к ячейкам.
     *
     * @author Vyatcheslav Suharnikov
     */
    _initialize: function () {
        var sheet = this._sheet;

        // Сначала удалим все, кроме 1 ячейки (ее не можем удалить).
        if (sheet.getMaxRows() > 1) {
            sheet.deleteRows(1, sheet.getMaxRows() - 1);
        }

        if (sheet.getMaxColumns() > 1) {
            sheet.deleteColumns(1, sheet.getMaxColumns() - 1);
        }

        // При создании новых ячеек, стиль будет копироваться из первой.
        sheet.getRange(1, 1)
            .setBackground('#efefef')
            .setVerticalAlignment('middle');

        // Добавим строк и столбцов столько, сколько нам нужно.
        var rows = 3            // Заголовок.
            + this._daysInMonth // Количество дней.
            + 1                 // Итого.
            + 6;                // Подвал.

        var cols = 16 // Основная таблица.
            + 4;  // Легенда.

        sheet.insertRowsBefore(1, rows - 1);
        sheet.insertColumnsBefore(1, cols - 1);
    },

    /**
     * Рисует заголовки основной таблицы со смещением (startColumn, startRow).
     *
     * @param {number} startRow    Смещение по строкам.
     * @param {number} startColumn Смещение по столбцам.
     *
     * @author Vyatcheslav Suharnikov
     */
    _setupHeader: function (startRow, startColumn) {
        var sheet = this._sheet;

        // Устанавливаем ширину колонок.
        var columnsWidth = [
            90,  // Дата.
            105, // День недели.

            // Проект 1.
            250, // Задачи
            50,  // Общее затраченное время.

            // Проект 2.
            250, // Задачи
            50,  // Общее затраченное время.

            // Проект 3.
            250, // Задачи
            50,  // Общее затраченное время.

            // Проект 4.
            250, // Задачи
            50,  // Общее затраченное время.

            50, // С.
            50, // По.
            50, // С.
            50, // По.
            50, // Часы.
            50, // D.
            50  // Суммарное время.
        ];

        for (var columnIndex = 0; columnIndex < columnsWidth.length; columnIndex++) {
            sheet.setColumnWidth(startColumn + columnIndex, columnsWidth[columnIndex]);
        }

        // Общие стили.
        sheet.getRange('A1:P3').offset(startRow - 1, startColumn - 1)
            .setFontWeight('bold')
            .setHorizontalAlignment('center')
            .setBorder(true, true, true, true, true, true);

        // Устанавливаем текст в заголовках.
        sheet.getRange(startRow, startColumn, 1, 3).setValues([
            ['Дата', 'День', 'Отчет о рабочем времени']
        ]);

        sheet.getRange(startRow, startColumn + 2).setValue('Отчет о рабочем времени %username%');

        sheet.getRange(startRow + 1, startColumn + 2).setValue('Задачи');
        sheet.getRange(startRow, startColumn + 10).setValue(this._reportName);

        sheet.getRange(startRow + 1, startColumn + 10, 1, 6).setValues([
            ['С', 'По', 'С', 'По', 'Часы', 'D']
        ]);

        // Разукрашиваем.
        var fillRanges = {
            'A1:P3': '#969696',
            'C1:P1': '#99ccff',
            'C3:J3': '#ffcc99'
        };

        for (var range in fillRanges) {
            sheet.getRange(range).offset(startRow - 1, startColumn - 1)
                .setBackgroundColor(fillRanges[range]);
        }

        // Устанавливаем необходимый размер шрифтов.
        var fontSizeRanges = {
            'A1:P3': 12,
            'C1': 14
        };

        for (var range in fontSizeRanges) {
            sheet.getRange(range).offset(startRow - 1, startColumn - 1)
                .setFontSize(fontSizeRanges[range]);
        }

        // Сливаем ячейки.
        var mergeRanges = [
            'C1:J1', // Отчет о …
            'K1:P1', // Месяц и год.
            'A1:A3', // Дата.
            'B1:B3', // День.
            'C2:J2', // Задачи.
            'C3:D3', // Проект 1.
            'E3:F3', // Проект 2.
            'G3:H3', // Проект 3.
            'I3:J3', // Проект 4.
            'K2:K3', // С.
            'L2:L3', // По.
            'M2:M3', // С
            'N2:N3', // По.
            'O2:O3', // Часы.
            'P2:P3' // D.
        ];

        for (var i in mergeRanges) {
            sheet.getRange(mergeRanges[i]).offset(startRow - 1, startColumn - 1)
                .merge();
        }
    },

    /**
     * Настройка строк для всех дней.
     *
     * @param {number} startRow     Стартовая строка.
     * @param {number} startColumn  Стартовый столбец.
     * @private
     */
    _setupDayRows: function (startRow, startColumn) {
        var sheet = this._sheet,
            currDayDate = this._date;

        for (var i = 0; i < this._daysInMonth; i++) {
            currDayDate.setDate(i + 1);
            this._setupDayRow(startRow + i, startColumn, currDayDate);
        }

        sheet.getRange(startRow, startColumn, this._daysInMonth, 16)
            .setBorder(true, true, true, true, true, true);
    },

    /**
     * Настройка строки для конкретного дня.
     *
     * @param {number} startRow     Стартовая строка.
     * @param {number} startColumn  Стартовый столбец.
     * @param {Date}   dayDate      Для какого дня настраиваем?
     * @private
     */
    _setupDayRow: function (startRow, startColumn, dayDate) {
        var sheet = this._sheet,
            backgroundColor = '#fff'; // Фон строки.

        if (reports.Utils.isHoliday(dayDate)) {
            backgroundColor = '#ccc';
        } else if (reports.Utils.isPreHoliday(dayDate)) {
            backgroundColor = '#efefef';
        }

        // Стили.
        sheet.getRange(startRow, startColumn, 1, 16)
            .setBackgroundColor(backgroundColor)
            .setFontWeight('normal')
            .setHorizontalAlignment('left');

        // Расположение текста в "С, По, С, По, Часы, D".
        sheet.getRange(startRow, startColumn + 10, 1, 6)
            .setHorizontalAlignment('right');

        // Текст, формулы и формат.
        sheet.getRange(startRow, startColumn)
            .setNumberFormat('dd.MM.YYYY')
            .setValue(dayDate);

        // Наименовение дня недели.
        sheet.getRange(startRow, startColumn + 1).setValue(reports.Utils.getDayName(dayDate));

        // Проект 1-4 - затраченное время.
        var cols = [3, 5, 7, 9];
        for (var i in cols) {
            sheet.getRange(startRow, startColumn + cols[i])
                .setNumberFormat('0.00');
        }

        // С, По, С, По.
        sheet.getRange(startRow, startColumn + 10, 1, 4)
            .setNumberFormat('00:00');

        // Часы.
        var firstPartFormula = 'HOUR(R[0]C[-1] - R[0]C[-2])+MINUTE(R[0]C[-1] - R[0]C[-2])/60',
            secondPartFormula = 'HOUR(R[0]C[-3] - R[0]C[-4])+MINUTE(R[0]C[-3] - R[0]C[-4])/60';
        sheet.getRange(startRow, startColumn + 14)
            .setNumberFormat('0.00')
            .setFormulaR1C1('=' + firstPartFormula + ' + ' + secondPartFormula);

        var expectedWorkTime = reports.Utils._getExpectedWorkHours(dayDate);

        // D.
        sheet.getRange(startRow, startColumn + 15)
            .setNumberFormat('0.00')
            .setFormulaR1C1('=R[0]C[-1] - ' + expectedWorkTime);
    },

    /**
     * Строка "Итого".
     *
     * @param {number} startRow     Стартовая строка.
     * @param {number} startColumn  Стартовый столбец.
     * @private
     */
    _setupSummaryRow: function (startRow, startColumn) {
        var sheet = this._sheet;

        // Стиль для ячеек по умолчанию.
        sheet.getRange(startRow, startColumn, 1, 16)
            .setFontWeight('bold')
            .setBackgroundColor('#969696')
            .setBorder(true, true, true, true, true, true);

        // Первые два столбца объединяем. 
        sheet.getRange(startRow, startColumn, 1, 2)
            .merge();

        // Суммарное затраченное время по каждому проекту.
        var columns = [3, 5, 7, 9]; // Индексы столбцов с ячейкой затраченного времени по отдельному проекту. 
        for (var i in columns) {
            // Считаем, что summary идет сразу после данных по дням.
            sheet.getRange(startRow, startColumn + columns[i])
                .setBackgroundColor('#ffcc99')
                .setNumberFormat('0.00')
                .setFormulaR1C1('=SUM(R[-' + (this._daysInMonth + 1) + ']C[0]:R[-1]C[0])');
        }

        // Суммарное затраченное время по всем проектам.
        sheet.getRange(startRow, startColumn + 10)
            .setBackgroundColor('#33cccc')
            .setNumberFormat('0.00')
            .setHorizontalAlignment('center')
            .setFormulaR1C1('=SUM(R[0]C[-10]:R[0]C[-1])');

        // Текст "Итого".
        sheet.getRange(startRow, startColumn + 11, 1, 3)
            .merge()
            .setFontColor('#fff')
            .setHorizontalAlignment('right')
            .setValue('Итого:');

        // Суммарное затраченное время.
        sheet.getRange(startRow, startColumn + 14)
            .setBackgroundColor('#33cccc')
            .setHorizontalAlignment('right')
            .setFormulaR1C1('=SUM(R[-' + (this._daysInMonth + 1) + ']C[0]:R[-1]C[0])');

        // Суммарная дельта. Сколько часов мы еще не отработали в этом месяце?
        var deltaCell = sheet.getRange(startRow, startColumn + 15)
            .setBackgroundColor('#33cccc')
            .setHorizontalAlignment('right')
            .setFormulaR1C1('=SUM(R[-' + (this._daysInMonth + 1) + ']C[0]:R[-1]C[0])');

        // Количество часов на этот месяц.
        sheet.getRange(startRow, startColumn + 16)
            .setBackgroundColor('#ccffcc')
            .setHorizontalAlignment('center')
            .setValue(-deltaCell.getValues()[0][0]); // Скопируем из дельты, только знак изменим на +.
    },

    /**
     * Настройка подвала.
     *
     * @param {number} startRow     Стартовая строка.
     * @param {number} startColumn  Стартовый столбец.
     * @private
     */
    _setupFooter: function (startRow, startColumn) {
        var sheet = this._sheet,
            helpText1 =
                'Данные в колонках Проект 1, Проект 2 …\n' +
                'заполняются в часах\n' +
                'например, "3 часа 15 минут" заполняется как "3,25"\n' +
                'при необходимости используются формулы "=3+15/60"',
            helpText2 =
                'Данные в прочих задачах\n' +
                'заполняются текстом\n' +
                'с указанием номеров задач',
            helpText3 =
                'Данные в колонках "С", "По"\n' +
                'заполняются в часах и минутах\n' +
                'например, "половина второго"\n' +
                'заполняется как "13:30"';

        // Вспомонательный текст 1.
        sheet.getRange(startRow + 1, startColumn, 4, 3)
            .merge()
            .setBackgroundColor('#fff')
            .setHorizontalAlignment('left')
            .setFontSize(10)
            .setValue(helpText1);

        // Вспомонательный текст 2.
        sheet.getRange(startRow + 1, startColumn + 8, 4, 2)
            .merge()
            .setBackgroundColor('#fff')
            .setHorizontalAlignment('left')
            .setFontSize(10)
            .setValue(helpText2);

        // Вспомонательный текст 3.
        sheet.getRange(startRow + 1, startColumn + 10, 4, 6)
            .merge()
            .setBackgroundColor('#fff')
            .setHorizontalAlignment('left')
            .setFontSize(10)
            .setValue(helpText3);
    },

    /**
     * Легенда - список задач, над которыми сотрудник работал в течении месяца.
     *
     * @param {number} startRow     Стартовая строка.
     * @param {number} startColumn  Стартовый столбец.
     * @private
     */
    _setupLegend: function (startRow, startColumn) {
        // Установим ширину столбцов.  
        var sheet = this._sheet,
            columnsWidth = [ // Ширина 1, 2 и третьей колонок.
                50,
                80,
                315
            ],
            text = 'В легенду вписываются все задачи,\n'
                + 'которые в отчете фигурируют только по ID',
            rowsInLegend = this._daysInMonth;

        for (var columnIndex = 0; columnIndex < columnsWidth.length; columnIndex++) {
            sheet.setColumnWidth(startColumn + columnIndex, columnsWidth[columnIndex]);
        }

        // Общие стили для строк в легенде.
        sheet.getRange(startRow + 1, startColumn, rowsInLegend + 1, 3)
            .setBackgroundColor('#fff');

        // Заголовок.
        sheet.getRange(startRow, startColumn + 1, 1, 2)
            .merge()
            .setFontWeight('bold')
            .setBackgroundColor('#969696')
            .setFontColor('#fff')
            .setHorizontalAlignment('center')
            .setValue('Легенда');

        // Количество часов.
        sheet.getRange(startRow + 1, startColumn, rowsInLegend, 1)
            .setNumberFormat('0.00')
            .setHorizontalAlignment('right')
            .setBorder(true, true, true, true, true, true);

        // Проект/Задача.
        sheet.getRange(startRow + 1, startColumn + 1, rowsInLegend, 1)
            .setHorizontalAlignment('center')
            .setFontWeight('bold')
            .setBorder(true, true, true, false, true, true);

        // Описание задачи.
        sheet.getRange(startRow + 1, startColumn + 2, rowsInLegend, 1)
            .setHorizontalAlignment('left')
            .setBorder(true, false, true, false, true, true);

        // Итого.
        sheet.getRange(startRow + rowsInLegend + 1, startColumn)
            .setFontWeight('bold')
            .setBackgroundColor('#99ccff')
            .setHorizontalAlignment('right')
            .setFormulaR1C1('=SUM(R[-' + rowsInLegend + ']C[0]:R[-1]C[0])');

        // Текст с описанием.
        sheet.getRange(startRow + rowsInLegend + 1, startColumn + 1, 2, 2)
            .merge()
            .setBackgroundColor('#fff')
            .setHorizontalAlignment('left')
            .setFontSize(10)
            .setValue(text);
    },

    /**
     * Настройка области для проверки введеных значений.
     *
     * @param startRow
     * @param startColumn
     * @param checkRangeFormula
     * @param checkLegendRangeFormula
     * @private
     */
    _setupCheckArea: function (startRow, startColumn, checkRangeFormula, checkLegendRangeFormula) {
        var sheet = this._sheet,
            text = 'Если в полях "Проверка" или "Проверка по легенде" указано не "0,00"\n'
                + 'это означает, что не все рабочее время\n'
                + 'расписано в проектах и прочих задачах (общей таблице)\n'
                + 'или, соответственно, в легенде (таблице справа).';

        // Общие стили.
        sheet.getRange(startRow, startColumn, 2, 2)
            .setBackgroundColor('#ff8080')
            .setFontWeight('bold')
            .setHorizontalAlignment('center');

        // Заголовки.
        sheet.getRange(startRow, startColumn, 1, 2)
            .setValues([['Проверка', 'Проверка по легенде']]);

        // Проверка.
        sheet.getRange(startRow + 1, startColumn)
            .setFormula(checkRangeFormula);

        // Проверка по легенде.
        sheet.getRange(startRow + 1, startColumn + 1)
            .setFormula(checkLegendRangeFormula);

        // Описание.
        sheet.getRange(startRow + 2, startColumn, 5, 2)
            .merge()
            .setBackgroundColor('#fff')
            .setFontSize(10)
            .setHorizontalAlignment('left')
            .setValue(text);
    }
};