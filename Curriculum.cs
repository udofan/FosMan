using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        [Description("Очная")]
        FullTime,               //очная
        [Description("Заочная")]
        PartTime,               //заочная
        [Description("Очно-заочная")]
        MixedTime,              //очно-заочная
        [Description("НЕИЗВЕСТНО")]
        Unknown    
    }

    /// <summary>
    /// Учебный план [kəˈrɪkjʊləm]
    /// </summary>
    internal class Curriculum {
        static Regex m_regexTestDirectionName = new(@"(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2})\s+(.*)$", RegexOptions.Compiled);
        static Regex m_regexParseSemester = new(@"Семестр\s+(\d)", RegexOptions.Compiled);
        //static bool m_programIsDetected = false;

        /// <summary>
        /// Направление подготовки
        /// </summary>
        public string DirectionName { get; set; }
        /// <summary>
        /// Код направления
        /// </summary>
        public string DirectionCode { get; set; }
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
        /// Уникальный ключ, по которому УП будут объединяться в группы
        /// </summary>
        public string GroupKey { get => $"{DirectionCode} - {DirectionName} - {Profile}"; }

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
                Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

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
                    var cellValue = row[colIdx] as string;//)?.ToUpper();
                    if (cellValue != null) {
                        if (!programIsDetected) {
                            var match = m_regexTestDirectionName.Match(cellValue);
                            if (match.Success) {
                                programIsDetected = true;
                                curriculum.DirectionCode = string.Join("", match.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries));
                                curriculum.DirectionName = match.Groups[2].Value.Trim();
                            }
                        }
                        if (cellValue.Contains("ПРОФИЛЬ", StringComparison.CurrentCultureIgnoreCase) || 
                            cellValue.Contains("Программа магистр", StringComparison.CurrentCultureIgnoreCase)) {
                            curriculum.Profile = getNextCellValue(row, colIdx);
                        }
                        if (cellValue.Contains("КАФЕДРА", StringComparison.CurrentCultureIgnoreCase)) {
                            curriculum.Department = getNextCellValue(row, colIdx);
                        }
                        if (cellValue.Contains("ФАКУЛЬТЕТ", StringComparison.CurrentCultureIgnoreCase)) {
                            curriculum.Faculty = getNextCellValue(row, colIdx);
                        }
                        if (cellValue.Contains("ФОРМА ОБУЧ", StringComparison.CurrentCultureIgnoreCase)) {
                            curriculum.FormOfStudy = DetectFormOfStudy(cellValue);
                            if (curriculum.FormOfStudy == EFormOfStudy.Unknown) {
                                curriculum.Errors.Add($"Не удалось определить форму обучения - {cellValue}");
                            }
                        }
                        if (cellValue.Contains("УЧЕБНЫЙ ГОД", StringComparison.CurrentCultureIgnoreCase)) {
                            curriculum.AcademicYear = getNextCellValue(row, colIdx);
                        }
                        if (cellValue.Contains("ОБРАЗОВАТЕЛЬНЫЙ СТАНД", StringComparison.CurrentCultureIgnoreCase)) {
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
                    form = EFormOfStudy.MixedTime;
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

        /// <summary>
        /// Найти дисциплину в УП или вернуть null
        /// </summary>
        /// <param name="disciplineName"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        internal CurriculumDiscipline FindDiscipline(string disciplineName) {
            return Disciplines.Values.FirstOrDefault(d => string.Compare(d.Name, disciplineName, true) == 0);
        }

        public override string ToString() {
            var name = $"{DirectionCode} {DirectionName} - {Profile} - Форма обучения: {FormOfStudy.GetDescription()}";
            return name;
        }
    }
}
