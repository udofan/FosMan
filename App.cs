using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Загруженные данные
    /// </summary>
    static internal class App {
        static Dictionary<string, Curriculum> m_curriculumDic = [];
        static Dictionary<string, Rpd> m_rpdDic = [];

        public static bool HasCurriculumFile(string fileName) => m_curriculumDic.ContainsKey(fileName);

        public static bool HasRpdFile(string fileName) => m_rpdDic.ContainsKey(fileName);

        static public bool AddCurriculum(Curriculum curriculum) {
            return m_curriculumDic.TryAdd(curriculum.SourceFileName, curriculum);
        }
        static public bool AddRpd(Rpd rpd) {
            return m_rpdDic.TryAdd(rpd.SourceFileName, rpd);
        }
    }
}
