using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    /// <summary>
    /// Работа с docx-шаблонами
    /// </summary>
    public static class Templates {
        public const string TEMPLATE_RPD_TABLE_EDU_WORK = "РПД_таблица_содержания_дисциплины.docx";
        public const string TEMPLATE_ABSTRACT_FOR_DISCIPLINE = "АННОТАЦИЯ_к_дисциплине.docx";
        public const string TEMPLATE_ABSTRACTS_FOR_RPD = "АННОТАЦИИ_к_РПД.docx";

        static Dictionary<string, string> m_files = [];

        static Templates() {
            var templateDir = Path.Combine(Environment.CurrentDirectory, FormMain.DIR_TEMPLATES);

            if (Directory.Exists(templateDir)) {
                m_files = Directory.GetFiles(templateDir, "*.docx").ToDictionary(f => Path.GetFileName(f), f => f);
                //m_templateFileNames = files.Select(f => Path.GetFileName(f)).Except(comboBoxRpdGenTemplates.Items.Cast<string>());
                //if (newTemplates.Any()) {
                //comboBoxRpdGenTemplates.Items.AddRange(newTemplates.ToArray());
                //}
            }
        }

        /// <summary>
        /// Элементы коллекции шаблонов
        /// </summary>
        public static Dictionary<string, string> Items => m_files;

        /// <summary>
        /// Адаптировать свойства EduWork под новую таблицу из шаблона
        /// </summary>
        /// <param name="eduWork"></param>
        /// <param name="workTable"></param>
        internal static void AdoptEduWorkForTemplateTable(EducationalWork eduWork, Table workTable, int topicCount) {
            //МЕТОД ДОЛЖЕН БЫТЬ доработан после изменения шаблона TEMPLATE_RPD_TABLE_EDU_WORK
            eduWork.TableTopicStartRow = 3;
            eduWork.TableHasContactTimeSubtotal = true;
            eduWork.TableColTopic = 0;
            eduWork.TableColEvalTools = 6;
            eduWork.TableColCompetenceResults = 7;
            eduWork.TableMaxColCount = 8;
            eduWork.TableTopicLastRow = eduWork.TableTopicStartRow + topicCount - 1;
            eduWork.TableControlRow = eduWork.TableTopicLastRow + 1;
            eduWork.TableStartNumCol = 1;
        }
    }
}
