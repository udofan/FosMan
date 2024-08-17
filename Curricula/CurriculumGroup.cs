using FastMember;
using System;
using System.Collections.Concurrent;
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
    /// Группа УП (исп. при генерации РПД)
    /// </summary>
    internal class CurriculumGroup : BaseObj {
        //static TypeAccessor m_typeAccessor = TypeAccessor.Create(typeof(CurriculumGroup));
        ConcurrentDictionary<string, CurriculumDiscipline> m_disciplines = [];
        ConcurrentDictionary<string, Curriculum> m_curricula = [];
        string m_formsOfStudyList = null;
        Department m_department = null;

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
        /// Кафедра
        /// </summary>
        public string DepartmentName { get; set; }
        /// <summary>
        /// Описание кафедры
        /// </summary>
        [JsonIgnore]
        public Department Department {
            get {
                if (m_department == null && !string.IsNullOrEmpty(DepartmentName)) {
                    m_department = App.Config.Departments.Values.FirstOrDefault(d => d.Name.Equals(DepartmentName, StringComparison.CurrentCultureIgnoreCase) ||
                                                                                     d.NameGenitive.Equals(DepartmentName, StringComparison.CurrentCultureIgnoreCase));
                }
                return m_department;
            }
        }
        /// <summary>
        /// Образовательный стандарт (ФГОС)
        /// </summary>
        public string FSES { get; set; }
        /// <summary>
        /// Формы обучения по всем УП из группы
        /// </summary>
        public List<EFormOfStudy> FormsOfStudy { get => Curricula?.Values.Select(c => c.FormOfStudy).ToList(); }
        /// <summary>
        /// Формы обучения в виде списка [исп. для вставки в РПД, Аннотации и т.п.]
        /// </summary>
        public string FormsOfStudyList {
            get {
                if (m_formsOfStudyList == null) {
                    var forms = FormsOfStudy.ToList();
                    forms.Sort((x1, x2) => (int)x1 - (int)x2);
                    m_formsOfStudyList = string.Join(", ", forms.Select(f => f.GetDescription())).ToLower();
                }
                return m_formsOfStudyList;
            }
            set => m_formsOfStudyList = value;
        }
        /// <summary>
        /// Квалификация
        /// </summary>
        public EDegree Degree { get => Curricula?.Values.FirstOrDefault().Degree ?? EDegree.Unknown; }
        /// <summary>
        /// Квалификация (для экрана)
        /// </summary>
        public string DegreeForScreen { get => Curricula?.Values.FirstOrDefault()?.DegreeForScreen ?? EDegree.Unknown.GetDescription(); }
        /// <summary>
        /// УП, входящие в группу
        /// </summary>
        public ConcurrentDictionary<string, Curriculum> Curricula { get => m_curricula; set => m_curricula = value; }
        /// <summary>
        /// Список дисциплин из группы
        /// </summary>
        public ConcurrentDictionary<string, CurriculumDiscipline> Disciplines { get => m_disciplines; }
        /// <summary>
        /// Дисциплины для генерации
        /// </summary>
        public List<CurriculumDiscipline> CheckedDisciplines { get; set; }

        public CurriculumGroup(Curriculum curriculum) {
            DepartmentName = curriculum.Department;
            DirectionCode = curriculum.DirectionCode;
            DirectionName = curriculum.DirectionName;
            Profile = curriculum.Profile;
            FSES = curriculum.FSES;
            
            AddCurriculum(curriculum);
            //Curricula = new[] { [curriculum.SourceFileName] = curriculum };
        }

        /// <summary>
        /// Добавить УП в группу
        /// </summary>
        /// <param name="curriculum"></param>
        /// <returns></returns>
        public bool AddCurriculum(Curriculum curriculum) {
            var result = false;
            
            if (m_curricula.TryAdd(curriculum.SourceFileName, curriculum)) {
                foreach (var disc in curriculum.Disciplines.Values) {
                    m_disciplines.TryAdd(disc.Key, disc);
                }

                result = true;
            }

            return true;
        }

        public override string ToString() {
            var name = $"{DirectionCode} {DirectionName} - {Profile} - Форм обучения: {FormsOfStudy?.Count ?? 0} - Планов в группе: {Curricula?.Count ?? 0}";
            return name;
        }

        /// <summary>
        /// Получить значение свойства по имени
        /// </summary>
        /// <param name="propName"></param>
        /// <returns></returns>
        //public object GetProperty(string propName) {
        //    object value = null;
        //    try {
        //        value = m_typeAccessor[this, propName];
        //    }
        //    catch (Exception ex) {
        //    }

        //    return value;
        //}
    }
}
