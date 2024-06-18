using Accessibility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Xceed.Words.NET;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Описание ФОСа
    /// </summary>
    internal class Fos : BaseObj {
        /// <summary>
        /// Ключ (по нему сравниваются РПД и ФОС)
        /// </summary>
        [JsonIgnore]
        public string Key { get => App.NormalizeName($"{DirectionCode}_{Profile}_{DisciplineName}"); }
        /// <summary>
        /// Исходный docx-файл
        /// </summary>
        [JsonInclude]
        public string SourceFileName { get; set; }
        /// <summary>
        /// Год
        /// </summary>
        [JsonInclude]
        public string Year { get; set; }
        /// <summary>
        /// Кафедра
        /// </summary>
        [JsonInclude]
        public string Department { get; set; }
        /// <summary>
        /// Составитель
        /// </summary>
        [JsonInclude]
        public string Compiler { get; set; }
        /// <summary>
        /// Дисциплина
        /// </summary>
        [JsonInclude]
        public string DisciplineName { get; set; }
        /// <summary>
        /// Формы обучения
        /// </summary>
        [JsonInclude]
        public List<EFormOfStudy> FormsOfStudy { get; set; }
        /// <summary>
        /// Дата утверждения
        /// </summary>
        //public DateTime ApprovalDate { get; set; }
        /// <summary>
        /// Код направления подготовки
        /// </summary>
        [JsonInclude]
        public string DirectionCode { get; set; }
        /// <summary>
        /// Направление подготовки
        /// </summary>
        [JsonInclude]
        public string DirectionName { get; set; }
        /// <summary>
        /// Профиль
        /// </summary>
        [JsonInclude]
        public string Profile { get; set; }
        /// <summary>
        /// Матрица компетенций (п.2 в ФОС)
        /// </summary>
        [JsonInclude]
        public CompetenceMatrix CompetenceMatrix { get; set; }
        /// <summary>
        /// Паспорт фонда оценочных средств текущего контроля, соотнесённых с индикаторами достижения компетенций (п.3)
        /// </summary>
        [JsonInclude]
        public List<StudyModule> Passport { get; set; }
        /// <summary>
        /// Выявленные ошибки
        /// </summary>
        [JsonIgnore]
        public List<string> Errors { get; set; }
        /// <summary>
        /// Список доп. ошибок, выявленных при проверке
        /// </summary>
        [JsonIgnore]
        public List<string> ExtraErrors { get; set; }

        /// <summary>
        /// Загрузка ФОС из файла
        /// </summary>
        /// <param name="fileName"></param>
        public static Fos LoadFromFile(string fileName, Fos fos = null) {
            fos ??= new();

            fos.Errors = [];
            fos.SourceFileName = fileName;
            fos.FormsOfStudy = [];
            //fos.SummaryParagraphs = [];
            //fos.QuestionList = [];
            //fos.ReferencesBase = [];
            //fos.ReferencesExtra = [];
            fos.CompetenceMatrix = null;
            fos.Passport = null;
            fos.Compiler = "";
            fos.Department = "";
            fos.DirectionCode = "";
            fos.DirectionName = "";
            fos.DisciplineName = "";
            fos.Profile = "";
            fos.Year = "";

            try {
                if (!DocParser.TryParse(fileName, fos, FosParser.Rules, out var errors)) {
                    fos.Errors.AddRange(errors);
                    return null;
                }
                else {
                    using (var docx = DocX.Load(fileName)) {
                        //цикл по таблицам
                        foreach (var table in docx.Tables) {
                            var keepTestTable = true;

                            //проверка на таблиц 2.1
                            if (keepTestTable &&
                                fos.CompetenceMatrix == null &&
                                CompetenceMatrix.TestTable(table, out var format) &&
                                format == ECompetenceMatrixFormat.Fos21) {
                                if (App.TestForTableOfCompetenceMatrix(table, format, out var matrix, out errors)) {
                                    fos.CompetenceMatrix = matrix;
                                    keepTestTable = false;
                                }
                                if (errors.Any()) fos.Errors.AddRange(errors.Select(e => $"Матрица компетенций: {e}"));
                            }
                            //проверка на таблицу 2.2
                            if (keepTestTable &&
                                fos.CompetenceMatrix != null &&
                                !fos.CompetenceMatrix.IsComplete &&
                                CompetenceMatrix.TestTable(table, out format) &&
                                format == ECompetenceMatrixFormat.Fos22) {
                                //попытка дополнить таблицу компетенций по второй таблице 2.2
                                if (CompetenceMatrix.TryParseTable(table, format, fos.CompetenceMatrix)) {
                                    fos.CompetenceMatrix.Check();
                                    if (fos.CompetenceMatrix.Errors.Any()) {
                                        fos.Errors.AddRange(fos.CompetenceMatrix.Errors.Select(e => $"Матрица компетенций: {e}"));
                                    }
                                    keepTestTable = false;
                                }
                            }
                            //проверка на таблицу "3. Паспорт фонда оценочных средств текущего контроля, соотнесённых с индикаторами достижения компетенций"
                            if (keepTestTable && fos.Passport == null) {
                                if (App.TestForFosPassport(table, out var passport, out errors)) {
                                    fos.Passport = passport;
                                    keepTestTable = false;
                                    if (errors.Any()) fos.Errors.AddRange(errors.Select(e => $"Паспорт: {e}"));
                                }
                            }
                        }
                        //итоговая проверка
                        if (fos.CompetenceMatrix == null || !fos.CompetenceMatrix.IsLoaded) {
                            fos.Errors.Add("Не найдена матрица компетенций.");
                        }
                        if (string.IsNullOrEmpty(fos.Department)) fos.Errors.Add("Не удалось определить название кафедры");
                        if (string.IsNullOrEmpty(fos.Profile)) fos.Errors.Add("Не удалось определить профиль");
                        if (string.IsNullOrEmpty(fos.Year)) fos.Errors.Add("Не удалось определить год программы");
                        if (string.IsNullOrEmpty(fos.DirectionCode)) fos.Errors.Add("Не удалось определить шифр направления подготовки");
                        if (string.IsNullOrEmpty(fos.DirectionName)) fos.Errors.Add("Не удалось определить наименование направления подготовки");
                        if (string.IsNullOrEmpty(fos.DisciplineName)) fos.Errors.Add("Не удалось определить название дисциплины");
                        if (fos.Passport == null) fos.Errors.Add("Не удалось определить паспорт");
                    }
                }
            }
            catch (Exception ex) {
                fos.Errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }

            return fos;
        }

    }
}
