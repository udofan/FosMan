using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Учебный план [kəˈrɪkjʊləm]
    /// </summary>
    internal class Curriculum : BaseObj {
        static Regex m_regexTestDirectionName = new(@"(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2})\s+(.*)$", RegexOptions.Compiled);
        static Regex m_regexParseSemester = new(@"Семестр\s+(\d)", RegexOptions.Compiled);
        //static bool m_programIsDetected = false;

        /// <summary>
        /// Направление подготовки
        /// </summary>
        [JsonInclude]
        public string DirectionName { get; set; }
        /// <summary>
        /// Код направления
        /// </summary>
        [JsonInclude]
        public string DirectionCode { get; set; }
        /// <summary>
        /// Профиль
        /// </summary>
        [JsonInclude]
        public string Profile { get; set; }
        /// <summary>
        /// Факультет
        /// </summary>
        [JsonInclude]
        public string Faculty { get; set; }
        /// <summary>
        /// Кафедра
        /// </summary>
        [JsonInclude]
        public string Department { get; set; }
        /// <summary>
        /// Форма обучения
        /// </summary>
        [JsonInclude]
        public EFormOfStudy FormOfStudy { get; set; } = EFormOfStudy.Unknown;
        /// <summary>
        /// Квалификация
        /// </summary>
        [JsonInclude]
        public EDegree Degree { get; set; }
        /// <summary>
        /// Квалификация (для экрана)
        /// </summary>
        public string DegreeForScreen { get => Degree.GetDescription().ToLower(); }
        /// <summary>
        /// Учебный год
        /// </summary>
        [JsonInclude]
        public string AcademicYear { get; set; }
        /// <summary>
        /// Образовательный стандарт (ФГОС)
        /// </summary>
        [JsonInclude]
        public string FSES { get; set; }
        /// <summary>
        /// Список описания дисциплин
        /// </summary>
        [JsonInclude]
        public Dictionary<string, CurriculumDiscipline> Disciplines { get; set; }
        /// <summary>
        /// Исходный xlsx-файл
        /// </summary>
        [JsonInclude]
        public string SourceFileName { get; set; }
        /// <summary>
        /// Ошибки
        /// </summary>
        [JsonIgnore]
        public List<string> Errors { get; set; }
        /// <summary>
        /// Уникальный ключ, по которому УП будут объединяться в группы
        /// </summary>
        [JsonIgnore]
        public string GroupKey { get => $"{DirectionCode} - {DirectionName} - {Profile}".ToLower(); }

        /// <summary>
        /// Загрузка описания УП из xlsx-файла
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static Curriculum LoadFromFile(string fileName, Curriculum curriculum) {
            curriculum ??= new();

            curriculum.SourceFileName = fileName;
            curriculum.Errors = [];
            curriculum.Disciplines = [];
            curriculum.FSES = string.Empty;
            curriculum.Faculty = string.Empty;
            curriculum.AcademicYear = string.Empty;
            curriculum.Degree = EDegree.Unknown;
            curriculum.Department = string.Empty;
            curriculum.DirectionCode = string.Empty;
            curriculum.DirectionName = string.Empty;
            curriculum.FormOfStudy = EFormOfStudy.Unknown;
            curriculum.Profile = string.Empty;

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
                        if (cellValue.Contains("Квалификация:", StringComparison.CurrentCultureIgnoreCase)) {
                            curriculum.Degree = DetectDegree(cellValue);
                        }
                    }
                }
            }
        }

        private static EDegree DetectDegree(string value) {
            var degree = EDegree.Unknown;

            value = value.Trim().ToLower();

            if (!string.IsNullOrEmpty(value)) {
                if (value.Contains("бакалавр")) {
                    degree = EDegree.Bachelor;
                }
                else if (value.Contains("магистр")) {
                    degree = EDegree.Master;
                }
            }

            return degree;
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
        internal CurriculumDiscipline FindDiscipline(string? disciplineName) {
            disciplineName = App.NormalizeName(disciplineName); // disciplineName.ToLower().Replace('ё', 'е');
            var discipline = Disciplines.Values.FirstOrDefault(d => App.NormalizeName(d.Name).Equals(disciplineName));
            discipline ??= Disciplines.Values.FirstOrDefault(d => App.NormalizeName(d.Name).StartsWith(disciplineName));
            discipline ??= Disciplines.Values.FirstOrDefault(d => disciplineName.StartsWith(App.NormalizeName(d.Name)));

            return discipline; // Disciplines.Values.FirstOrDefault(d => d.Name.ToLower().Replace('ё', 'е').Equals(disciplineName));
        }

        public override string ToString() {
            var name = $"{DirectionCode} {DirectionName} - {Profile} - Форма обучения: {FormOfStudy.GetDescription()}";
            return name;
        }

        public bool AddDiscipline(CurriculumDiscipline discipline) {
            var result = Disciplines.TryAdd(discipline.Key, discipline);
            if (result) {
                discipline.Curriculum = this;
            }
            return result;
        }
    }
}
