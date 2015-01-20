/**
 * Регистрация пунктов меню в текущей таблице.
 *
 * @param {Object} options Настройки.
 *
 * @option options {string} menuTitle   Заголовок в меню.
 * @option options {string} importName  Наименование библиотеки, указанное при подключении.
 */
function Register(options) {
    reports.general.ensureReport();

    // Почему функции находятся в "reports."?
    // Потому что это - глобальный объект (указывается при настройке), в который загружаются скрипты из проекта.
    SpreadsheetApp.getUi()
        .createMenu(options.menuTitle)
        .addItem('Общий - Скрыть дни', options.importName + '.General_hideDays')
        .addItem('Общий - Скрыть минуты', options.importName + '.General_hideMinutes')
        .addSeparator()
        .addItem('Ежемесячный - Текущий', options.importName + '.Monthly_createForCurrent')
        .addItem('Ежемесячный - Следующий', options.importName + '.Monthly_createForNext')
        .addToUi();
}

// В виду ограничений функции addItem.

function General_hideMinutes() {
    reports.general.hideMinutes();
}

function General_hideDays() {
    reports.general.hideDays();
}

function Monthly_createForCurrent() {
    var report = new reports.Monthly(reports.Utils.getCurrentMonthFirstDayDate());
    report.makeEmptyReport();

    reports.general.ensureFirst();
}

function Monthly_createForNext() {
    var report = new reports.Monthly(reports.Utils.getNextMonthFirstDayDate());
    report.makeEmptyReport();

    reports.general.ensureFirst();
}