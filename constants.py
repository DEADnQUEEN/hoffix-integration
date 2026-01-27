HOFFIX_LOGIN_URL = "https://hoffix.hoff.ru/login"
HOFFIX_MAIN_URL = "https://hoffix.hoff.ru/orders"
CONFIRM_LOGIN_SELECTOR = ".el-button.el-button--primary"
LOGIN_PASSWORD_FIELDS_SELECTOR = ".el-input__inner"

STEP = 300

WORKER_SHEET = "Список мастеров"
SERVICES_SHEET = "Services"

EXCEL_DATE_FORMAT = "%d.%m.%Y"
HOFFIX_DATE_FORMAT = "%Y-%m-%d"
OUTPUT_DATE_FORMAT = "%d-%m-%YT%H:%M:%S"

WORKER_SELECT = "form[class='el-form grid-table'] .grid-table__row:nth-child(2) div:has(> input.el-input__inner)"
POSSIBLE_SELECT_OPTIONS = "body > .el-select-dropdown.el-popper .el-scrollbar__view.el-select-dropdown__list > .el-select-dropdown__item[style=''], body > .el-select-dropdown.el-popper .el-scrollbar__view.el-select-dropdown__list > .el-select-dropdown__item:not([style])"
TABLE_ELEMENTS = ".el-table__body > tbody > tr.el-table__row"

EDIT_BUTTON = "#orderBlock > .data-card__header .el-button.el-button--default.el-button--medium:has(span)"
SAVE_BUTTON = "#orderBlock > .data-card__header .el-button.el-button--primary.el-button--medium:has(span)"
WAIT_SCRIPT = "() => {return document.querySelectorAll('#orderBlock > .data-card__header .el-button.el-button--primary.el-button--medium.is-loading:has(span)').length === 0;}"

OUTPUT_LIST = "Протокол"
SETTING_LIST = "Настройки"
LOGIN_PASSWORD_FIELD = "B9"

ATTEMPTS = 5