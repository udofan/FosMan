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
                Errors = []
            };

            try {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                var programIsDetected = false;

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
                        var tableTitle = dataSet.Tables["Титул"];
                        if (tableTitle == null) {
                            curriculum.Errors.Add($"Не найден лист \"Титул\"");
                        }
                        else {
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
                    }
                }
            }
            catch (Exception ex) {
                curriculum.Errors.Add(ex.Message);
            }

            return curriculum;
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
