using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Описание кафедры
    /// </summary>
    internal class Department : BaseObj {
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
                        Name = "",
                        NameGenitive = "",
                        Compiler = "",
                        Signer = ""
                    },
                    ["3"] = new Department {
                        Code = "3",
                        Name = "",
                        NameGenitive = "",
                        Compiler = "",
                        Signer = ""
                    },
                    ["4"] = new Department {
                        Code = "4",
                        Name = "",
                        NameGenitive = "",
                        Compiler = "",
                        Signer = ""
                    },
                    ["5"] = new Department {
                        Code = "5",
                        Name = "",
                        NameGenitive = "",
                        Compiler = "",
                        Signer = ""
                    },
                    ["6"] = new Department {
                        Code = "6",
                        Name = "Государственное администрирование",
                        NameGenitive = "государственного администрирования",
                        Compiler = "",
                        Signer = ""
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
                        Name = "",
                        NameGenitive = "",
                        Compiler = "",
                        Signer = ""
                    },
                    ["9"] = new Department {
                        Code = "9",
                        Name = "",
                        NameGenitive = "",
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
