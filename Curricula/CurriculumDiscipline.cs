﻿using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Описание дисциплины
    /// </summary>
    public class CurriculumDiscipline : BaseObj {
        public const int SEMESTER_COUNT = 10;

        static Regex m_regexTestTypeRequired = new(@"^[^\.]+\.О\.\d+", RegexOptions.Compiled);      //обязательная
        static Regex m_regexTestTypeVariable = new(@"^[^\.]+\.В\.\d+", RegexOptions.Compiled);      //обязательная часть, формируемая участниками
        static Regex m_regexTestTypeByChoice = new(@"^[^\.]+\.В\.ДВ\.\d+", RegexOptions.Compiled);  //выбору части, формируемой...
        static Regex m_regexTestTypeOptional = new(@"^ФТД", RegexOptions.Compiled);
        static Regex m_regexBlockNum = new(@"Б(\d+)\.", RegexOptions.Compiled);
        HashSet<string> m_competenceList = null;
        EducationalWork m_eduWork = null;
        Department m_department = null;

        string m_competences = null;
        EDisciplineType? m_type = null;
        int m_startSemIdx = -1;
        int m_lastSemIdx = -1;
        int m_blockNum = -1;

        /// <summary>
        /// Наименование дисциплины [ключ]
        /// </summary>
        [JsonInclude]
        public string Name { get; set; }
        [JsonIgnore]
        public string Key { get => App.NormalizeName(Name); }
        /// <summary>
        /// Порядковый номер дисциплины (= номер ряда в таблице Плана УП)
        /// </summary>
        [JsonInclude]
        public int Number { get; set; }
        /// <summary>
        /// Индекс дисциплины
        /// </summary>
        [JsonInclude]
        public string Index { get; set; }
        /// <summary>
        /// Номер блока обучения (вынимается из индекса)
        /// </summary>
        [JsonInclude]
        public int BlockNum {
            get {
                if (m_blockNum <= 0) {
                    var m = m_regexBlockNum.Match(Index);
                    if (m.Success) {
                        m_blockNum = int.Parse(m.Groups[1].Value);
                    }
                }
                return m_blockNum;
            }
        }
        /// <summary>
        /// Тип дисциплины
        /// </summary>
        [JsonIgnore]
        public EDisciplineType? Type {
            get {
                if (!string.IsNullOrEmpty(Index) && m_type == null) {
                    if (m_regexTestTypeOptional.IsMatch(Index)) {
                        m_type = EDisciplineType.Optional;
                    }
                    else if (m_regexTestTypeVariable.IsMatch(Index)) {
                        m_type = EDisciplineType.Variable;
                    }
                    else if (m_regexTestTypeByChoice.IsMatch(Index)) {
                        m_type = EDisciplineType.ByChoice;
                    }
                    else if (m_regexTestTypeRequired.IsMatch(Index)) {
                        m_type = EDisciplineType.Required;
                    }
                    else {
                        m_type = EDisciplineType.Unknown;
                    }
                }
                return m_type;
            }
        }
        /// <summary>
        /// Описание типа дисциплины [поле для РПД]
        /// </summary>
        [JsonIgnore]
        public string TypeDescription {
            get {
                var text = "";
                var degree = Curriculum?.Degree == Enums.EDegree.Master ? "магистратуры" : "бакалавриата";

                if (Type == EDisciplineType.Required) {
                    text = $"обязательным дисциплинам программы {degree}";
                }
                else if (Type == EDisciplineType.Variable) {
                    text = $"обязательным дисциплинам части, формируемой участниками образовательных отношений программы {degree}";
                }
                else if (Type == EDisciplineType.ByChoice) {
                    text = $"дисциплинам по выбору части, формируемой участниками образовательных отношений программы {degree}";
                }
                else if (Type == EDisciplineType.Optional) {
                    text = $"программе {degree}, формируемой участниками образовательных отношений";
                }
                return text;
            }
        }

        /// <summary>
        /// Доп. описание типа дисциплины [поле для РПД]
        /// </summary>
        [JsonIgnore]
        public string TypeDescription2 {
            get {
                var text = "";
                if (Type == EDisciplineType.Optional) {
                    text = " и относится к факультативной дисциплине";
                }
                
                return text;
            }
        }

        /// <summary>
        /// Тип контроля в регистре: зачет, экзамен
        /// </summary>
        [JsonIgnore]
        public string ControlType { get => EducationalWork?.ControlFormForScreen.ToLower() ?? ""; }

        /// <summary>
        /// Название закрепленной кафедры
        /// </summary>
        [JsonInclude]
        public string DepartmentName { get; set; }
        /// <summary>
        /// Код закрепленной кафедры
        /// </summary>
        [JsonInclude]
        public string DepartmentCode { get; set; }
        /// <summary>
        /// Описание кафедры
        /// </summary>
        [JsonIgnore]
        public Department Department { 
            get {
                if (m_department == null && !string.IsNullOrEmpty(DepartmentCode)) {
                    App.Config.Departments.TryGetValue(DepartmentCode, out m_department);
                }
                return m_department;
            }
        }
        /// <summary>
        /// Нормализованные компетенции (коды)
        /// </summary>
        [JsonIgnore]
        public HashSet<string> CompetenceList {
            get {
                if (m_competenceList == null && !string.IsNullOrEmpty(Competences)) {
                    var competenceItems = Competences.Split(';', StringSplitOptions.TrimEntries);
                    m_competenceList = competenceItems.Select(x => string.Join("", x.Split(' ', StringSplitOptions.TrimEntries)).ToUpper()).ToHashSet();
                }
                return m_competenceList;
            }
        }
        /// <summary>
        /// Компетенции строкой (берутся из xlsx)
        /// </summary>
        [JsonInclude]
        public string Competences { get; set; }
        /// <summary>
        /// Формы контроля
        /// </summary>
        //public List<EControlForm> ControlForms { get; set; }
        [JsonInclude]
        public string? ControlFormExamSemester { get; set; }
        [JsonInclude]
        public string? ControlFormTestSemester { get; set; }
        [JsonInclude]
        public string? ControlFormTestWithAGradeSemester { get; set; }
        [JsonInclude]
        public string? ControlFormCourseWorkSemester { get; set; }
        /// <summary>
        /// Итого акад.часов: по плану
        /// </summary>
        [JsonInclude]
        public int? TotalByPlanHours { get; set; }
        /// <summary>
        /// Итого акад.часов: Конт. раб.
        /// </summary>
        [JsonInclude]
        public int? TotalContactWorkHours { get; set; }
        /// <summary>
        /// Итого акад.часов: СР
        /// </summary>
        [JsonInclude]
        public int? TotalSelfStudyHours { get; set; }
        /// <summary>
        /// Итого акад.часов: Конт роль
        /// </summary>
        [JsonInclude]
        public int? TotalControlHours { get; set; }
        /// <summary>
        /// Описание учебных работ по семестрам
        /// </summary>
        [JsonInclude]
        public EducationalWork[] Semesters { get; set; } = new EducationalWork[SEMESTER_COUNT];

        /// <summary>
        /// Кол-во зачетных единиц (1 ЗЕ = 36 часов)
        /// </summary>
        [JsonIgnore]
        public int? TestUnits { get => TotalByPlanHours / 36; }

        /// <summary>
        /// Выявленные ошибки
        /// </summary>
        [JsonIgnore]
        public List<string> Errors { get; set; }
        /// <summary>
        /// Доп. ошибки (исп. при проверке по загруженной матрице компетенций)
        /// </summary>
        [JsonIgnore]
        public List<string> ExtraErrors { get; set; }
        /// <summary>
        /// Родительский УП
        /// </summary>
        [JsonIgnore]
        public Curriculum Curriculum { get; set; }

        /// <summary>
        /// Объект описания учебной работы
        /// </summary>
        [JsonIgnore]
        public EducationalWork EducationalWork {
            get {
                m_eduWork ??= new() {
                    ContactWorkHours = TotalContactWorkHours,
                    ControlForm = EControlForm.Unknown,
                    ControlHours = TotalControlHours,
                    LabHours = Semesters.Sum(s => s.LabHours ?? 0),
                    LectureHours = Semesters.Sum(s => s.LectureHours ?? 0),
                    PracticalHours = Semesters.Sum(s => s.PracticalHours ?? 0),
                    SelfStudyHours = TotalSelfStudyHours,
                    TotalHours = TotalByPlanHours
                };
                if (!string.IsNullOrEmpty(ControlFormExamSemester)) {
                    m_eduWork.ControlForm = EControlForm.Exam;
                }
                else if (!string.IsNullOrEmpty(ControlFormTestSemester)) {
                    m_eduWork.ControlForm = EControlForm.Test;
                }
                else if (!string.IsNullOrEmpty(ControlFormTestWithAGradeSemester)) {
                    m_eduWork.ControlForm = EControlForm.TestWithAGrade;
                }
                else if (!string.IsNullOrEmpty(ControlFormCourseWorkSemester)) {
                    m_eduWork.ControlForm = EControlForm.CourseWork;
                }
                
                return m_eduWork;
            }
        }

        /// <summary>
        /// Индекс стартового семестра
        /// </summary>
        [JsonIgnore]
        public int StartSemesterIdx {
            get {
                if (m_startSemIdx < 0) {
                    var sem = this.Semesters.FirstOrDefault(s => s.TotalHours.HasValue && s.TotalHours > 0);
                    if (sem != null) {
                        m_startSemIdx = Array.IndexOf(this.Semesters, sem);
                    }
                }
                return m_startSemIdx;
            }
        }

        /// <summary>
        /// Индекс последнего семестра
        /// </summary>
        [JsonIgnore]
        public int LastSemesterIdx {
            get {
                if (m_lastSemIdx < 0) {
                    var sem = this.Semesters.LastOrDefault(s => s.TotalHours.HasValue && s.TotalHours > 0);
                    if (sem != null) {
                        m_lastSemIdx = Array.IndexOf(this.Semesters, sem);
                    }
                }
                return m_lastSemIdx;
            }
        }

        public CurriculumDiscipline() {
            Errors = [];
            ExtraErrors = [];

            for (var i = 0; i < Semesters.Length; i++) {
                Semesters[i] = new();
            }
        }

        /// <summary>
        /// Проверка дисциплины
        /// </summary>
        /// <param name="curriculum"></param>
        internal bool Check(Curriculum curriculum, int rowIdx) {
            var result = false;

            if (string.IsNullOrEmpty(Name)) {
                return result;
            }

            if (Type == EDisciplineType.Unknown) {
                var err = $"Не удалось определить тип дисциплины [{Name}] по индексу [{Index}] (строка {rowIdx})";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            //проверка итоговых час ов
            var total = (TotalContactWorkHours ?? 0) + (TotalControlHours ?? 0) + (TotalSelfStudyHours ?? 0);
            if (TotalByPlanHours != total) {
                var err = $"Дисциплина [{Name}] (строка {rowIdx}): обнаружено неверное значение в поле [По плану]";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            //проверка часов семестров
            var allSemTotal = 0;
            var allSemContact = 0;
            for (var i = 0; i < Semesters.Length; i++) {
                var sem = Semesters[i];
                var semTotal = (sem.LectureHours ?? 0) + (sem.LabHours ?? 0) + (sem.SelfStudyHours ?? 0) + (sem.PracticalHours ?? 0) + (sem.ControlHours ?? 0);
                if (sem.TotalHours != semTotal) {
                    var err = $"Дисциплина [{Name}] (строка {rowIdx}): обнаружено неверное значение в поле [Итого] Семестра {i + 1}";
                    curriculum.Errors.Add(err);
                    Errors.Add(err);
                }
                allSemTotal += semTotal;
                var semContact = (sem.LectureHours ?? 0) + (sem.LabHours ?? 0) + (sem.PracticalHours ?? 0);
                allSemContact += semContact;
            }
            if (TotalByPlanHours != allSemTotal) {
                var err = $"Дисциплина [{Name}] (строка {rowIdx}): сумма итогового времени по семестрам не сходится с полем [Итого]";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            if (TotalContactWorkHours != allSemContact) {
                var err = $"Дисциплина [{Name}] (строка {rowIdx}): сумма контактного времени по семестрам не сходится с полем [Конт. раб.]";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            //if (TotalContactWorkHours != prac)
            //пропуска строку с названием модуля - у нее не заполнена ячейка кафедры
            if (!string.IsNullOrEmpty(DepartmentName)) {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// Проверка компетенций по указанной матрице
        /// </summary>
        /// <returns></returns>
        internal bool CheckCompetences(CompetenceMatrix matrix) {
            ExtraErrors = [];

            if (matrix == null) {
                ExtraErrors.Add("Матрица компетенций не загружена. Проверить компетенции не удалось.");
                return false;
            }

            var achiCodeList = matrix.GetAllAchievementCodes();
            if (CompetenceList == null || CompetenceList.Count == 0) {
                ExtraErrors.Add($"Список компетенций не определён.");
            }
            else {
                foreach (var achiCode in CompetenceList) {
                    if (!achiCodeList.Contains(achiCode)) {
                        ExtraErrors.Add($"В загруженной матрице компетенций не найден индикатор достижений [{achiCode}].");
                    }
                }
            }

            return ExtraErrors.Count == 0;
        }
    }
}
