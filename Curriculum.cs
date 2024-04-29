using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Форма обучения
    /// </summary>
    public enum EFormOfStudy {
        FullTime,               //очная
        PartTime,               //заочная
        FullTimeAndPartTime,    //очно-заочная
        Unknown    
    }

    /// <summary>
    /// Учебный план [kəˈrɪkjʊləm]
    /// </summary>
    internal class Curriculum {
        static Regex m_regexTestProgramName = new(@"(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2})\s+(.*)$", RegexOptions.Compiled);
        static Regex m_regexParseSemester = new(@"Семестр\s+(\d)", RegexOptions.Compiled);
        //static bool m_programIsDetected = false;

        /// <summary>
        /// Название программы
        /// </summary>
        public string ProgramName { get; set; }
        /// <summary>
        /// Код программы
        /// </summary>
        public string ProgramCode { get; set; }
        /// <summary>
        /// Профиль
        /// </summary>
        public string Profile { get; set; }
        /// <summary>
        /// Факультет
        /// </summary>
        public string Faculty { get; set; }
        /// <summary>
        /// Кафедра
        /// </summary>
        public string Department { get; set; }
        /// <summary>
        /// Форма обучения
        /// </summary>
        public EFormOfStudy FormOfStudy { get; set; } = EFormOfStudy.Unknown;
        /// <summary>
        /// Учебный год
        /// </summary>
        public string AcademicYear { get; set; }
        /// <summary>
        /// Образовательный стандарт (ФГОС)
        /// </summary>
        public string FSES { get; set; }
        /// <summary>
        /// Список описания дисциплин
        /// </summary>
        public Dictionary<string, CurriculumDiscipline> Disciplines { get; set; }
        /// <summary>
        /// Исходный xlsx-файл
        /// </summary>
        public string SourceFileName { get; set; }
        /// <summary>
        /// Ошибки
        /// </summary>
        public List<string> Errors { get; set; }

        /// <summary>
        /// Загрузка описания УП из xlsx-файла
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static Curriculum LoadFromFile(string fileName) {
            var curriculum = new Curriculum() {
                SourceFileName = fileName,
                Errors = [], 
                Disciplines = []
            };

            try {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var fileStream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite)) {
                    using (var reader = ExcelReaderFactory.CreateReader(fileStream)) {
                        var config = new ExcelDataSetConfiguration() {
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration() {
                                UseHeaderRow = true,
                                ReadHeaderRow = (rowHeader) => rowHeader.Read(),
                                FilterColumn = (rowHeader, columnIndex) => true,
                                FilterRow = (rowHeader) => true
                            }
                        };

                        DataSet dataSet = reader.AsDataSet(config);

                        ParseTableTitle(curriculum, dataSet);
                        ParseTablePlan(curriculum, dataSet);
                    }
                }
            }
            catch (Exception ex) {
                curriculum.Errors.Add(ex.Message);
            }

            return curriculum;
        }

        /// <summary>
        /// Парсинг листа "Титул"
        /// </summary>
        /// <param name="curriculum"></param>
        /// <param name="dataSet"></param>
        static void ParseTableTitle(Curriculum curriculum, DataSet dataSet) {
            var tableTitle = dataSet.Tables["Титул"];
            if (tableTitle == null) {
                curriculum.Errors.Add($"Не найден лист \"Титул\"");
                return;
            }

            //вернуть значение следующей значащей ячейки
            Func<DataRow, int, string> getNextCellValue = (row, colIdx) => {
                string value = null;

                for (int col = colIdx + 1; col < tableTitle.Columns.Count; col++) {
                    value = row[col] as string;
                    if (!string.IsNullOrEmpty(value)) {
                        break;
                    }
                }

                return value?.Trim();
            };

            var programIsDetected = false;

            //обработка Титула
            for (int rowIdx = 0; rowIdx < tableTitle.Rows.Count; rowIdx++) {
                var row = tableTitle.Rows[rowIdx];
                for (int colIdx = 0; colIdx < tableTitle.Columns.Count; colIdx++) {
                    var cellValue = (row[colIdx] as string)?.ToUpper();
                    if (cellValue != null) {
                        if (!programIsDetected) {
                            var match = m_regexTestProgramName.Match(cellValue);
                            if (match.Success) {
                                programIsDetected = true;
                                curriculum.ProgramCode = string.Join("", match.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries));
                                curriculum.ProgramName = match.Groups[2].Value.Trim();
                            }
                        }
                        if (cellValue.Contains("ПРОФИЛЬ")) {
                            curriculum.Profile = getNextCellValue(row, colIdx);
                        }
                        if (cellValue.Contains("КАФЕДРА")) {
                            curriculum.Department = getNextCellValue(row, colIdx);
                        }
                        if (cellValue.Contains("ФАКУЛЬТЕТ")) {
                            curriculum.Faculty = getNextCellValue(row, colIdx);
                        }
                        if (cellValue.Contains("ФОРМА ОБУЧ")) {
                            curriculum.FormOfStudy = DetectFormOfStudy(cellValue);
                            if (curriculum.FormOfStudy == EFormOfStudy.Unknown) {
                                curriculum.Errors.Add($"Не удалось определить форму обучения - {cellValue}");
                            }
                        }
                        if (cellValue.Contains("УЧЕБНЫЙ ГОД")) {
                            curriculum.AcademicYear = getNextCellValue(row, colIdx);
                        }
                        if (cellValue.Contains("ОБРАЗОВАТЕЛЬНЫЙ СТАНД")) {
                            curriculum.FSES = getNextCellValue(row, colIdx);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Парсинг листа "План"
        /// </summary>
        /// <param name="curriculum"></param>
        /// <param name="dataSet"></param>
        static void ParseTablePlan(Curriculum curriculum, DataSet dataSet) {
            try {
                var tablePlan = dataSet.Tables["План"];
                if (tablePlan == null) {
                    curriculum.Errors.Add($"Не найден лист \"План\"");
                    return;
                }

                //обработка Плана
                Dictionary<int, CurriculumDisciplineHeader> headers = null;

                for (int rowIdx = 0; rowIdx < tablePlan.Rows.Count; rowIdx++) {
                    var row = tablePlan.Rows[rowIdx];
                    if (headers?.Any() ?? false) {
                        CurriculumDisciplineReader.ProcessRow(tablePlan, rowIdx, headers, curriculum);
                    }
                    else {
                        CurriculumDisciplineReader.TestHeaderRow(tablePlan, rowIdx, curriculum, out headers);
                    }
                    /*
                    if (valuedRow) {
                        discipline = new() {
                            CompetenceList = []
                        };
                        if (colIndex >= 0) {
                            discipline.Index = row[colIndex] as string;
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [Индекс]");
                        }
                        if (colName >= 0) {
                            discipline.Name = row[colName] as string;
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [Наименование] дисциплины");
                        }
                        if (colDepartmentCode >= 0) {
                            discipline.DepartmentCode = row[colDepartmentCode] as string;
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [Код] закрепленной кафедры");
                        }
                        if (colDepartment >= 0) {
                            discipline.Department = row[colDepartment] as string;
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [Наименование] закрепленной кафедры");
                        }
                        if (colControlFormExam >= 0) {
                            if (int.TryParse(row[colControlFormExam] as string, out int val)) discipline.ControlFormExamHours = val;
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [Экзамен]");
                        }
                        if (colControlFormTest >= 0) {
                            if (int.TryParse(row[colControlFormTest] as string, out int val)) discipline.ControlFormTestHours = val;
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [Зачет]");
                        }
                        if (colControlFormTestWithAGrade >= 0) {
                            if (int.TryParse(row[colControlFormTestWithAGrade] as string, out int val)) discipline.ControlFormTestWithAGradeHours = val;
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [Зачет с оц.]");
                        }
                        if (colControlFormControlWork >= 0) {
                            if (int.TryParse(row[colControlFormControlWork] as string, out int val)) discipline.ControlFormControlWorkHours = val;
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [КР]");
                        }

                        if (colCompetenceList >= 0) {
                            var competenceList = row[colCompetenceList] as string;
                            if (!string.IsNullOrEmpty(competenceList)) {
                                var competenceItems = competenceList.Split(';', StringSplitOptions.TrimEntries);
                                discipline.CompetenceList.AddRange(competenceItems.Select(x => string.Join("", x.Split(' '))));
                            }
                        }
                        else {
                            curriculum.Errors.Add($"Не удалось обнаружить колонку [Компетенции]");
                        }
                        if (!curriculum.Disciplines.TryAdd(discipline.Name, discipline)) {
                            curriculum.Errors.Add($"Повторное упоминание дисциплины [{discipline.Name}] на строке {rowIdx}");
                        }
                    }
                    else {
                        for (int colIdx = 0; colIdx < tablePlan.Columns.Count; colIdx++) {
                            var cellValue = (row[colIdx] as string)?.ToUpper();
                            if (cellValue != null) {
                                //попытка парсинга ряда с номерами семестров
                                if (!semesterRowIsParsed) {
                                    //семестр X
                                    var match = m_regexParseSemester.Match(cellValue);
                                    if (match.Success && int.TryParse(match.Groups[1].Value.Trim(), out int num)) {
                                        semesterColIdx[num] = colIdx;
                                        semesterRow = true;
                                    }
                                }
                                //попытка парсинга ряда заголовков
                                if (!headerRowIsParsed) {
                                    if (cellValue == "ИНДЕКС") {
                                        colIndex = colIdx;
                                        headerRow = true;
                                    }
                                    if (cellValue == "НАИМЕНОВАНИЕ") {
                                        if (colIdx == colIndex + 1) {
                                            colName = colIdx;
                                        }
                                        else {
                                            colDepartment = colIdx;
                                        }
                                    }
                                    if (cellValue.Contains("ЭКЗА")) colControlFormExam = colIdx;
                                    if (cellValue.Contains("ЗАЧЕТ С ОЦ")) {
                                        colControlFormTestWithAGrade = colIdx;
                                    }
                                    else if (cellValue.Contains("ЗАЧЕТ")) {
                                        colControlFormTest = colIdx;
                                    }
                                    if (cellValue.Contains("КОД")) {
                                        colDepartmentCode = colIdx;
                                    }
                                    if (cellValue.Contains("КОМПЕТЕНЦИИ")) {
                                        colCompetenceList = colIdx;
                                    }
                                    if (cellValue == "КР") {
                                        colControlFormControlWork = colIdx;
                                    }
                                }
                                else {
                                    if (valuedRow) { //значащий ряд

                                    }
                                }
                            }
                        }
                        if (semesterRow) semesterRowIsParsed = true;
                        if (headerRow) {
                            headerRowIsParsed = true;
                        }
                    }
                    */
                }
            }
            catch (Exception ex) {
                curriculum.Errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Определение формы обучения
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        static EFormOfStudy DetectFormOfStudy(string value) {
            var form = EFormOfStudy.Unknown;

            value = value.Trim().ToLower();

            if (!string.IsNullOrEmpty(value)) {
                if (value.Contains("очно-заочная")) {
                    form = EFormOfStudy.FullTimeAndPartTime;
                }
                else if (value.Contains("заочная")) {
                    form = EFormOfStudy.PartTime;
                }
                else if (value.Contains("очная")) {
                    form = EFormOfStudy.FullTime;
                }
            }

            return form;
        }
    }
}
