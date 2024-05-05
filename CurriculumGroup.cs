using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Группа УП (исп. при генерации РПД)
    /// </summary>
    internal class CurriculumGroup {
        Dictionary<string, CurriculumDiscipline> m_disciplines = [];
        Dictionary<string, Curriculum> m_curricula = [];

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
        public string Department { get; set; }
        /// <summary>
        /// Формы обучения по всем УП из группы
        /// </summary>
        public List<EFormOfStudy> FormsOfStudy { get => Curricula?.Values.Select(c => c.FormOfStudy).ToList(); }
        /// <summary>
        /// УП, входящие в группу
        /// </summary>
        public Dictionary<string, Curriculum> Curricula { get => m_curricula; }
        public Dictionary<string, CurriculumDiscipline> Disciplines { get => m_disciplines; }
        /// <summary>
        /// Дисциплины для генерации
        /// </summary>
        public List<CurriculumDiscipline> CheckedDisciplines { get; set; }

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
    }
}
