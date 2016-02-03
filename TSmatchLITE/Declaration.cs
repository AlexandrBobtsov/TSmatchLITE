using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSmatch.Declaration
{
    /// <summary>
    /// Declaration -- определения глобальных констант
    /// </summary>
    /// <journal> 20.12.2013
    /// 7.1.2014    - добавлена секция констант Шаблонов
    /// 23.1.14     - переход к DataTable
    /// 12.12.15    - адаптация к TSmatch
    /// 22.12.15    - секция типов Штампа Документа
    /// 2.1.2016    - моды работы с Документом
    /// 9.1.2016    - разделитель списков в строке ";"
    /// 24.1.16     - константы Matching_Rules, регулярнае выражения
    /// </journal>
    class Declaration
    {
        //------------ КОНСТАНТЫ ОБЩЕГОНАЗНАЧЕНИЯ -----------------
        public const char STR_DELIMITER = ';';  // знак - резделитель в списке в строках
        /// <summary>
        /// F_MATCH = "TSmatch.xlsx" - имя файла таблиц приложения TSmatch
        /// !!временно разместил вместе с Tekla Structures\версия\Environments, но надо будет уточнить в Tekla!!
        /// </summary>
        public const string F_MATCH = "TSmatch.xlsx";
        public const string DIR_MATCH = @"C:\ProgramData\Tekla Structures\21.1\Environments\common\exceldesign";
        public const string DOC_TOC = "TOC";    // Table Of Content - Лист - таблица-содержание всех Документов в TSmatch

        //--------- листы TSmatch.xlsm --------------
        public const string CONST = "Const";
        public const string HEADER = "Header";
        public const string LOG = "Log";

        //-----------константы таблицы Документов -----------------
        public const int DOC_TIME = 1; // дата и время последнего изменения Документа
        public const int DOC_NAME = 2; // имя Документа
        public const int DOC_EOL = 3; // EOL Документа
        public const int DOC_MYCOL = 4; // количество колонок, дабавляемых слева в Документ в loadDoc
        public const int DOC_RESLINES = 5; // число строк пятки
        public const int DOC_MADESTEP = 6; // последний выполненный Шаг
        public const int DOC_DIR = 7; // каталог, где лежит Документ
        public const int DOC_FILE = 8; // файл match, содержащий Документ
        public const int DOC_SHEET = 9; // лист Документа
        public const int DOC_STMPTXT = 10; // текст Штампа
        public const int DOC_STMPTYPE = 11; // тип Штампа
        public const int DOC_STMPROW = 12; // строка Штампа
        public const int DOC_STMPCOL = 13; // колонка Штампа
        public const int DOC_CREATED = 14; // дата создания Документа
        public const int DOC_PATTERN = 16; // основной шаблон Документа. Помимо заголовков Шапки,
                                           //..он содержит описания колонок данных
        public const int DOC_SUMMARY_PATTERN = 17; // шаблон пятки
        public const int DOC_LOADER = 20; // Loader Документа
        public const int DOC_MD5 = 27; // Check Sum MD5

        //-----------константы моды работы с Документом -----------------
        public enum DOC_RW_TYPE { RW, RO, CREATEOROPEN, N1, N2, N };
        public const int READWRITE = 0;   // Мода Документа - открыть на чтение и запись -- по умолчанию
        public const int READONLY = 3;   // Мода Документа - только чтение
        public const int APPEND = 11;   // Мода Документа - дописывать
        public const int CREATEOROPEN = 21;   // Мода Документа - если не найден - создать

        //----------- ТИПЫ ШТАМПОВ / ТИПЫ ДОКУМЕНТОВ ----------------------
        public const string STAMP_TYPE_EQ = "=";  // точное соответствие
        public const string STAMP_TYPE_INC = "I"; // Includes
        public const string STAMP_TYPE_N = "N";   // New Doc без Штампа или No Stamp
        public const string STAMP_TYPE_N1 = "N1"; // New Doc, создавать новый Лист в самой левой позиции, если его нет; для Summary
        public const string STAMP_TYPE_N2 = "N2"; // New Doc, если нет, создавать в Листе2

        //-----------константы таблицы Процессов -----------------
        public const string PROCESS = "Process";

        public const int STEP_TIME = 0;    // Дата и время выполнения Шага. Если тут пусто или "/" - это строка-комментарий
        public const int PROC_NAME = 1;    // Имя Процесса. В сроках Шага тут пусто
        public const int STEP_PREVSTEP = 2;    // Предшествующий Шаг. Шаг этой строки выполняем только
                                               //..если предыдущий Шаг выполнен или тут пусто
        public const int STEP_NAME = 3;    // Имя Шага
        public const int STEP_DONE = 4;    // Шаг выполнен "1", иначе - выполняем
        public const int STEP_PARAM = 5;    // Параметры Шага
        public const int STEP_INPDOCS = 6;    // Входные Документы Шага
        public const int STEP_OUTDOCS = 7;    // Выходные Документы Шага

        //------------ константы журнала моделей -------------------
        public const string MODELS = "Models";

        public const int MODEL_DATE = 1;    // Дата и время записи в Журнал моделей
        public const int MODEL_NAME = 2;    // Имя модели
        public const int MODEL_DIR = 3;    // Каталог модели
        public const int MODEL_MADE = 4;    // Модель обработана Шагом (Made)
        public const int MODEL_MD5 = 5;    // Контрольная сумма - MD5

        //------------ константы правил Matching_Rules -------------------
        public const string MATCHING_RULES = "Matching_Rules";

        public const int RULE_DATE = 1;    // Дата и время записи Правила в TSmatch
        public const int RULE_NAME = 2;    // Имя Правила
        public const int RULE_TYPE = 3;    // Тип Правила
        public const int RULE_PARS = 4;    // Параметры Правила
        public const int RULE_DOCS = 5;    // Документы Правила

        public const string ATT_DELIM = @"(${must}|,|=| |\t|\*|x|X|х|Х)";  //делимитры в Правилах
        public const string ATT_MUST  = @"(?<must>(&'.+')|(&"".+"")|(&«.+»)|(&“.+”)"; // &обязательно
        public const string ATT_PARAM = @"(?<param>(\$|p|р|п|P|Р|П)\w*\d)"; //параметры в Правилах

        //------------ Файл TSmatchINFO.xlsx - записывается в каталог модели ----------
        public const string SH1_MODELINFO = "ModelINFO"; //Лист общей информации по модели
        public const string SH2_RAW = "Raw";        //Лист необработанных данных по компонентам модели
        public const string SH3_REPORT = "Report";     //Лист - отчет по модели

        public const string TS_PATH = "#TS.ModInfo.ModePath";   //#template в TOC- вместо ссылки
                                                                //на каталог модели

        //-----------константы Шаблонов -----------------
        public const string PTRN_HDR = "A1";     // заголовки колонок   
        public const string PTRN_WIDTH = "A3";     // ширина колонок
        public const int PTRN_FETCH = 6;    // Fetch запрос

        public const string PTRN_COPYHDR = "CopyHdr"; // указание копировать заголовок из Шаблона
    }
}