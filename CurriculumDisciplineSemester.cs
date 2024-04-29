using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Описание семестра
    /// </summary>
    internal class CurriculumDisciplineSemester {
        public int? TotalHours { get; set; } = 0;
        public int? LectureHours { get; set; } = 0;
        public int? LabHours { get; set; } = 0;
        public int? PracticalHours { get; set; } = 0;
        public int? SelfStudyHours { get; set; } = 0;
        public int? ControlHours { get; set; } = 0;
    }
}
