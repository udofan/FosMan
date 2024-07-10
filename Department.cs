using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Описание кафедры
    /// </summary>
    public class Department : BaseObj {
        /// <summary>
        /// Название
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Название в родительном падеже
        /// </summary>
        public string NameGenitive { get; set; }
        /// <summary>
        /// Код кафедры
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// Составитель (полное наименование)
        /// </summary>
        public string Compiler { get; set; }
        /// <summary>
        /// Подписант (рассмотрена и принята)
        /// </summary>
        public string Signer { get; set; }


        /// <summary>
        /// Дефолтный список кафедр
        /// </summary>
        public static Dictionary<string, Department> DefaultDepartments {
            get {
                Dictionary<string, Department> dic = new() {
                    ["1"] = new Department {
                        Code = "1",
                        Name = "Социально-гуманитарные и естественнонаучные дисциплины",
                        NameGenitive = "социально-гуманитарных и естественнонаучных дисциплин",
                        Compiler = "",
                        Signer = ""
                    },
                    ["2"] = new Department {
                        Code = "2",
                        Name = "Теория и история государства и права",
                        NameGenitive = "теории и истории государства и права",
                        Compiler = "",
                        Signer = ""
                    },
                    ["3"] = new Department {
                        Code = "3",
                        Name = "Государственно-правовые дисциплины, административное и муниципальное право",
                        NameGenitive = "государственно-правовых дисциплин, административного и муниципального права",
                        Compiler = "",
                        Signer = ""
                    },
                    ["4"] = new Department {
                        Code = "4",
                        Name = "Гражданско-правовые дисциплины и международное частное право",
                        NameGenitive = "гражданско-правовых дисциплин и международного частного права",
                        Compiler = "",
                        Signer = ""
                    },
                    ["5"] = new Department {
                        Code = "5",
                        Name = "Уголовно-правовые дисциплины",
                        NameGenitive = "уголовно-правовых дисциплин",
                        Compiler = "",
                        Signer = ""
                    },
                    ["6"] = new Department {
                        Code = "6",
                        Name = "Государственное администрирование",
                        NameGenitive = "государственного администрирования",
                        Compiler = "к.э.н., профессор Скрынченко Б.Л.",
                        Signer = "Б.Л. Скрынченко"
                    },
                    ["7"] = new Department { 
                        Code = "7",
                        Name = "Экономика и менеджмент", 
                        NameGenitive = "экономики и менеджмента", 
                        Compiler = "к.э.н., Красавина В.А.", 
                        Signer = "В.А. Красавина"
                    },
                    ["8"] = new Department {
                        Code = "8",
                        Name = "Психология и педагогика",
                        NameGenitive = "психологии и педагогики",
                        Compiler = "",
                        Signer = ""
                    },
                    ["9"] = new Department {
                        Code = "9",
                        Name = "Педагогическое образование, специальная психология и дефектология",
                        NameGenitive = "педагогического образования, специальной психологии и дефектологии",
                        Compiler = "",
                        Signer = ""
                    },
                    ["10"] = new Department {
                        Code = "10",
                        Name = "Математика и информационные технологии",
                        NameGenitive = "математики и информационных технологий",
                        Compiler = "",
                        Signer = ""
                    }
                };
                return dic;
            }
        }
    }
}
