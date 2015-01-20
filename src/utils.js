// Вспомогательные функции: праздники, написание дат на русском языке и т.д..
// В пространстве имен reports.utils.
if (typeof reports === 'undefined') {
    reports = {};
}

/**
 * Вспомогательные утилиты.
 *
 * @constructor
 */
var Utils = function () {
};

/**

 Чтобы получить значения для очередного года:

 1. Зайдите на страницу производственного календаря сайта superjob.ru (для других не получится),
 например, http://www.superjob.ru/proizvodstvennyj_kalendar/ .
 2. Откройте консоль Google Chrome: cmd + shift + i
 3. Выполните указанный ниже код:

 var holidays = { holiday: [], pre: [] };
 $('.pk_container').each(function(i, monthElement) {
            var $monthElement = $(monthElement);
            
            holidays.holiday[i + 1] = $.map($monthElement.find('.pk_holiday'), function (dayElement) { return +dayElement.textContent.toString(); });
            holidays.pre[i + 1] = $.map($monthElement.find('.pk_preholiday'), function (dayElement) { return +dayElement.textContent.toString(); });
        });
 JSON.stringify(holidays);

 4. Скопировать получившийся результат без обрамляющих кавычек ({…})
 5. Вставить в _holidays с соотв. годом, например:

 this._holidays = { 2079: {…} };

 */
Utils._holidays = {
    2014: {
        "holiday": [null, [1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 18, 19, 25, 26], [1, 2, 8, 9, 15, 16, 22, 23], [1, 2, 8, 9, 10, 15, 16, 22, 23, 29, 30], [5, 6, 12, 13, 19, 20, 26, 27], [1, 2, 3, 4, 9, 10, 11, 17, 18, 24, 25, 31], [1, 7, 8, 12, 13, 14, 15, 21, 22, 28, 29], [5, 6, 12, 13, 19, 20, 26, 27], [2, 3, 9, 10, 16, 17, 23, 24, 30, 31], [6, 7, 13, 14, 20, 21, 27, 28], [4, 5, 11, 12, 18, 19, 25, 26], [1, 2, 3, 4, 8, 9, 15, 16, 22, 23, 29, 30], [6, 7, 13, 14, 20, 21, 27, 28]],
        "pre": [null, [], [24], [7], [30], [8], [11], [], [], [], [], [], [31]]
    },
    2015: {
        "holiday": [null, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 17, 18, 24, 25, 31], [1, 7, 8, 14, 15, 21, 22, 23, 28], [1, 7, 8, 9, 14, 15, 21, 22, 28, 29], [4, 5, 11, 12, 18, 19, 25, 26], [1, 2, 3, 4, 9, 10, 11, 16, 17, 23, 24, 30, 31], [6, 7, 12, 13, 14, 20, 21, 27, 28], [4, 5, 11, 12, 18, 19, 25, 26], [1, 2, 8, 9, 15, 16, 22, 23, 29, 30], [5, 6, 12, 13, 19, 20, 26, 27], [3, 4, 10, 11, 17, 18, 24, 25, 31], [1, 4, 7, 8, 14, 15, 21, 22, 28, 29], [5, 6, 12, 13, 19, 20, 26, 27]],
        "pre": [null, [], [], [], [30], [8], [11], [], [], [], [], [3], [31]]
    }
};

// Наименование дней на русском языке.
Utils._dayNames = [
    'Воскресение', // "Я вижу слезы ватников" (С) Ванга.
    'Понедельник',
    'Вторник',
    'Среда',
    'Четверг',
    'Пятница',
    'Суббота'
];

Utils.monthNames = [
    'Январь',
    'Февраль',
    'Март',
    'Апрель',
    'Май',
    'Июнь',
    'Июль',
    'Август',
    'Сентябрь',
    'Октябрь',
    'Ноябрь',
    'Декабрь'
];

Utils.prototype = {
    /**
     * Праздник?
     *
     * @param   {Date} date
     * @returns {boolean}
     */
    isHoliday: function (date) {
        var year = date.getFullYear(),
            month = date.getMonth() + 1,
            day = date.getDate(),
            dayOfWeek = date.getDay(),
            holidays = this._getHolidaysInfoForYear(year).holiday;

        return dayOfWeek === 0 // Воскресение.
            || dayOfWeek === 6 // Суббота.
            || holidays[month] && holidays[month].indexOf(day) >= 0; // Или есть в списке выходных.
    },

    /**
     * Предпраздничный день?
     *
     * @param   {Date} date
     * @returns {boolean}
     */
    isPreHoliday: function (date) {
        var year = date.getFullYear(),
            month = date.getMonth() + 1,
            day = date.getDate(),
            pre = this._getHolidaysInfoForYear(year).pre;

        return pre[month] && pre.indexOf(day) >= 0;
    },

    /**
     * Возвращает наименование дня на русском языке.
     * В Google App Scripts на данный момент (02.12.2014) не реализовано, даже со сменой локали!
     *
     * @param {Date} date
     * @returns {string}
     */
    getDayName: function (date) {
        return Utils._dayNames[date.getDay()];
    },

    /**
     * Возвращает первый день текущего месяца.
     * При форматировании возможны приколы с часовыми поясами, гугл переодически лихорадит.
     *
     * @returns {Date}
     */
    getCurrentMonthFirstDayDate: function () {
        var result = new Date();
        result.setDate(1);
        result.setHours(15);   // Возможно придется сменить.
        result.setMinutes(59);
        result.setSeconds(59);

        return result;
    },

    /**
     * Возвращает первый день для следующего месяца.
     * Возможны приколы с часовыми поясами, гугл переодически лихорадит.
     *
     * @returns {Date}
     */
    getNextMonthFirstDayDate: function () {
        var currDate = new Date();

        // Добавим 15:59:59, иначе форматирование будет тупить из-за приколов с часовыми поясами.
        return new Date(currDate.getFullYear(), currDate.getMonth() + 1, 1, 15, 59, 59);
    },

    /**
     * Количество дней в месяце.
     *
     * @param   {Date}  date В каком таком месяце?
     * @returns {number}
     */
    getDaysInMonth: function (date) {
        var lastDayOfMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0);
        return lastDayOfMonth.getDate();
    },

    /**
     * Форматирует дату в виде "Декабрь 2015".
     *
     * Факин гугл не умеет месяца по-русски (02.12.2014). Не веришь (а мб сделали)? Попробуй:
     * SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetLocale('RU_ru');
     * return Utilities.formatDate(date, 'GMT', 'MMMM Y');
     *
     * @param   {Date} date
     * @returns {string}
     */
    formatDateMMMMY: function (date) {
        return Utils.monthNames[date.getMonth()] + ' ' + date.getYear();
    },

    /**
     * Проверяет и в случае необходимости перемещает вкладку (по имени) на указанную позицию.
     *
     * @param {string} sheetName    Наименование вкладки.
     * @param {number} position     Позиция, которую необходимо установить для вкладки. >= 1.
     */
    ensureSheetNameIndex: function (sheetName, position) {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

        if (sheet) {
            this.ensureSheetIndex(sheet, position);
        }
    },

    /**
     * Проверяет и в случае необходимости перемещает вкладку на указанную позицию.
     *
     * @param {Sheet}   sheet       Вкладка.
     * @param {number}  position    Позиция, которую необходимо установить для вкладки. >= 1.
     */
    ensureSheetIndex: function (sheet, position) {
        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

        spreadsheet.setActiveSheet(sheet);
        spreadsheet.moveActiveSheet(position);
    },

    /**
     * Информация по праздникам за указанный год.
     *
     * @param   {number} year
     * @returns {*} Объект с двумя свойствами: holiday (список праздничных дней) и pre (список предпраздничных дней).
     * @private
     */
    _getHolidaysInfoForYear: function (year) {
        if (!Utils._holidays[year]) {
            throw new Error('Данные по праздникам за ' + year + ' не введены, обратитесь к документации в reports.Utils.');
        }

        return Utils._holidays[year];
    },

    /**
     * Возвращает ожидаемое количество часов работы в указанный день.
     *
     * @param   {Date} date
     * @return  {number}
     *
     * @private
     */
    _getExpectedWorkHours: function (date) {
        var result = 8; // Значение по умолчанию.

        if (this.isHoliday(date)) {
            result = 0; // В выходные не работаем.
        } else if (this.isPreHoliday(date)) {
            result -= 1; // В предпраздничные дни рабочий день сокращается на 1.
        }

        return result;
    }
};

reports.Utils = new Utils();