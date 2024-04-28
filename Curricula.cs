using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Учебные планы
    /// </summary>
    static internal class Curricula {
        static Dictionary<string, Curriculum> m_curriculumDic = new();

        public static bool HasFile(string fileName) => m_curriculumDic.ContainsKey(fileName);

        static public bool AddCurriculum(Curriculum curriculum) {
            return m_curriculumDic.TryAdd(curriculum.SourceFileName, curriculum);
        }
    }
}
