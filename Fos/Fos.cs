using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Описание ФОСа
    /// </summary>
    internal class Fos : BaseObj {
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
        public List<FosPassport> Passport { get; set; }
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
                }
            }
            catch (Exception ex) {

            }

            return fos;
        }
    }
}
