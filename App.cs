using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

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

        /// <summary>
        /// Extension для ячейки: получение текста по всем абзацам, объединенными заданной строкой
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="joinParText"></param>
        /// <returns></returns>
        static public string GetText(this Cell cell, string joinParText = " ", bool applyTrim = true) {
            var text = string.Join(joinParText, cell.Paragraphs.Select(p => p.Text));
            if (applyTrim) {
                return text.Trim();
            }
            return text;
        }
    }
}
