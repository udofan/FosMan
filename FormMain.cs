using BrightIdeasSoftware;
using Microsoft.VisualBasic;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Schema;
using Xceed.Document.NET;
using static FosMan.App;
using static FosMan.Enums;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace FosMan {
    public partial class FormMain : Form {
        public const string DIR_TEMPLATES = "Templates";
        public const string DIR_REPORTS = "Reports";
        public const string DIR_LOGS = "Logs";

        string m_matrixFileName = null;

        public FormMain() {
            InitializeComponent();
        }

        /// <summary>
        /// Покажем матрицу компетенций в WebView2
        /// </summary>
        /// <param name="errors"></param>
        async void ShowMatrix(CompetenceMatrix matrix) {
            var html = matrix.CreateHtmlReport();

            await webView21.EnsureCoreWebView2Async();
            webView21.NavigateToString(html);
        }

        void TuneCurriculumList() {
            var olvColumnCurriculumIndex = new OLVColumn("Код", "DirectionCode") {
                Width = 80,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumIndex);
            var olvColumnCurriculumName = new OLVColumn("Направление подготовки", "DirectionName") {
                Width = 200,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumName);
            var olvColumnCurriculumProfile = new OLVColumn("Профиль", "Profile") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumProfile);
            var olvColumnCurriculumDepartment = new OLVColumn("Кафедра", "Department") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumDepartment);
            var olvColumnCurriculumFaculty = new OLVColumn("Факультет", "Faculty") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumFaculty);
            var olvColumnCurriculumAcademicYear = new OLVColumn("Учебный год", "AcademicYear") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumAcademicYear);
            var olvColumnCurriculumFormOfStudy = new OLVColumn("Форма обучения", "FormOfStudy") {
                Width = 100,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var curriculum = x as Curriculum;
                    if (curriculum != null) {
                        if (curriculum.FormOfStudy == Enums.EFormOfStudy.FullTime) {
                            value = Enums.EFormOfStudy.FullTime.GetDescription();
                        }
                        else if (curriculum.FormOfStudy == Enums.EFormOfStudy.PartTime) {
                            value = Enums.EFormOfStudy.PartTime.GetDescription();
                        }
                        else if (curriculum.FormOfStudy == Enums.EFormOfStudy.MixedTime) {
                            value = Enums.EFormOfStudy.MixedTime.GetDescription();
                        }
                        else {
                            value = Enums.EFormOfStudy.Unknown.GetDescription();
                        }
                    }
                    return value;
                }
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumFormOfStudy);
            var olvColumnCurriculumErrors = new OLVColumn("Ошибки", "Errors") {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var curriculum = x as Curriculum;
                    if (curriculum != null) {
                        if (curriculum.Errors?.Any() ?? false) {
                            value = "есть";
                        }
                        else {
                            value = "нет";
                        }
                    }
                    return value;
                }
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumErrors);
            var olvColumnCurriculumFileName = new OLVColumn("Файл", "SourceFileName") {
                Width = 200,
                IsEditable = false,
                FillsFreeSpace = true
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumFileName);

            fastObjectListViewCurricula.PrimarySortColumn = olvColumnCurriculumProfile;
        }

        void TuneDisciplineList(FastObjectListView list, bool addSemesterColumns) {
            var olvColumnIndex = new OLVColumn("Индекс", "Index") {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnIndex);
            var olvColumnName = new OLVColumn("Наименование", "Name") {
                Width = 300,
                IsEditable = false
            };
            list.Columns.Add(olvColumnName);
            var olvColumnType = new OLVColumn("Тип", "Type") {
                Width = 100,
                IsEditable = false,
                AspectGetter = (x) => {
                    var discipline = x as CurriculumDiscipline;
                    var value = discipline?.Type?.GetDescription();
                    return value;
                }
            };
            list.Columns.Add(olvColumnType);
            var olvColumnErrors = new OLVColumn("Ошибки", "Errors") {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var discipline = x as CurriculumDiscipline;
                    if (discipline != null) {
                        if ((discipline.Errors?.Any() ?? false) || (discipline.ExtraErrors?.Any() ?? false)) {
                            value = "есть";
                        }
                        else {
                            value = "нет";
                        }
                    }
                    return value;
                }
            };
            list.Columns.Add(olvColumnErrors);

            var olvColumnTotalByPlanHours = new OLVColumn("Итого: По плану", "TotalByPlanHours") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnTotalByPlanHours);
            var olvColumnTotalContactWorkHours = new OLVColumn("Итого: Конт. раб.", "TotalContactWorkHours") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnTotalContactWorkHours);
            var olvColumnTotalSelfStudyHours = new OLVColumn("Итого: СР", "TotalSelfStudyHours") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnTotalSelfStudyHours);
            var olvColumnTotalControlHours = new OLVColumn("Итого: КР", "TotalControlHours") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnTotalControlHours);
            var olvColumnDepartmentName = new OLVColumn("Кафедра", "DepartmentName") {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDepartmentName);
            var colWidth = 55;
            if (addSemesterColumns) {
                for (var i = 0; i < CurriculumDiscipline.SEMESTER_COUNT; i++) {
                    var colIdx = i;
                    var semTitle = $"Сем. {i + 1}: ";
                    var olvColumnSemTotal = new OLVColumn($"{semTitle}Итого", "") {
                        Width = colWidth,
                        IsEditable = false,
                        AspectGetter = (x) => {
                            var value = "";
                            var discipline = x as CurriculumDiscipline;
                            if (discipline != null) {
                                value = discipline.Semesters[colIdx].TotalHours?.ToString() ?? "";
                            }
                            return value;
                        }
                    };
                    var olvColumnSemLec = new OLVColumn($"{semTitle}Лек", "") {
                        Width = colWidth,
                        IsEditable = false,
                        AspectGetter = (x) => {
                            var value = "";
                            var discipline = x as CurriculumDiscipline;
                            if (discipline != null) {
                                value = discipline.Semesters[colIdx].LectureHours?.ToString() ?? "";
                            }
                            return value;
                        }
                    };
                    list.Columns.Add(olvColumnSemLec);
                    var olvColumnSemLab = new OLVColumn($"{semTitle}Лаб", "") {
                        Width = colWidth,
                        IsEditable = false,
                        AspectGetter = (x) => {
                            var value = "";
                            var discipline = x as CurriculumDiscipline;
                            if (discipline != null) {
                                value = discipline.Semesters[colIdx].LabHours?.ToString() ?? "";
                            }
                            return value;
                        }
                    };
                    list.Columns.Add(olvColumnSemLab);
                    var olvColumnSemPractical = new OLVColumn($"{semTitle}Пр", "") {
                        Width = colWidth,
                        IsEditable = false,
                        AspectGetter = (x) => {
                            var value = "";
                            var discipline = x as CurriculumDiscipline;
                            if (discipline != null) {
                                value = discipline.Semesters[colIdx].PracticalHours?.ToString() ?? "";
                            }
                            return value;
                        }
                    };
                    list.Columns.Add(olvColumnSemPractical);
                    var olvColumnSemSelfStudy = new OLVColumn($"{semTitle}СР", "") {
                        Width = colWidth,
                        IsEditable = false,
                        AspectGetter = (x) => {
                            var value = "";
                            var discipline = x as CurriculumDiscipline;
                            if (discipline != null) {
                                value = discipline.Semesters[colIdx].SelfStudyHours?.ToString() ?? "";
                            }
                            return value;
                        }
                    };
                    list.Columns.Add(olvColumnSemSelfStudy);
                    var olvColumnSemControl = new OLVColumn($"{semTitle}Контроль", "") {
                        Width = colWidth,
                        IsEditable = false,
                        AspectGetter = (x) => {
                            var value = "";
                            var discipline = x as CurriculumDiscipline;
                            if (discipline != null) {
                                value = discipline.Semesters[colIdx].ControlHours?.ToString() ?? "";
                            }
                            return value;
                        }
                    };
                    list.Columns.Add(olvColumnSemControl);
                }
            }
            var olvColumnTotalCompetenceList = new OLVColumn("Компетенции", "CompetenceList") {
                Width = 300,
                IsEditable = false,
                //FillsFreeSpace = true,
                AspectGetter = x => {
                    var value = x?.ToString();
                    var discipline = x as CurriculumDiscipline;
                    if (discipline != null && discipline.CompetenceList != null) {
                        value = string.Join("; ", discipline.CompetenceList);
                    }
                    return value;
                }
            };
            list.Columns.Add(olvColumnTotalCompetenceList);

            list.PrimarySortColumn = olvColumnName;
        }

        /// <summary>
        /// Настройка списка РПД
        /// </summary>
        /// <param name="list"></param>
        void TuneRpdList(FastObjectListView list) {
            var olvColumnDisciplineName = new OLVColumn("Дисциплина", nameof(Rpd.DisciplineName)) {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDisciplineName);
            var olvColumnDirectionCode = new OLVColumn("Код", nameof(Rpd.DirectionCode)) {
                Width = 80,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDirectionCode);
            var olvColumnDirectionName = new OLVColumn("Направление подготовки", nameof(Rpd.DirectionName)) {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDirectionName);
            var olvColumnProfile = new OLVColumn("Профиль", nameof(Rpd.Profile)) {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnProfile);
            var olvColumnDepartment = new OLVColumn("Кафедра", nameof(Rpd.Department)) {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDepartment);
            var olvColumnYear = new OLVColumn("Год", nameof(Rpd.Year)) {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnYear);
            var olvColumnFormsOfStudy = new OLVColumn("Формы обучения", nameof(Rpd.FormsOfStudy)) {
                Width = 120,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var rpd = x as Rpd;
                    if (rpd != null) {
                        var items = rpd.FormsOfStudy.Select(x => x.GetAttribute<EvaluationToolAttribute>()?.ShortDescription ?? x.GetDescription()).ToList();
                        value = string.Join(", ", items);
                    }
                    return value;
                }
            };
            list.Columns.Add(olvColumnFormsOfStudy);
            //составитель
            var olvColumnCompiler = new OLVColumn("Составитель", nameof(Rpd.Compiler)) {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnCompiler);
            //матрица компетенций
            var olvColumnCompetenceMatrix = new OLVColumn("Компетенции", nameof(Rpd.CompetenceMatrix)) {
                Width = 60,
                CheckBoxes = true,
                IsEditable = false,
                AspectGetter = (x) => (x as Rpd)?.CompetenceMatrix?.Items?.Any() ?? false
            };
            list.Columns.Add(olvColumnCompetenceMatrix);
            //объем дисциплины (часы по формам обучения)
            var olvColumnEducationWork = new OLVColumn("Объем", nameof(Rpd.EducationalWorks)) {
                Width = 60,
                CheckBoxes = true,
                IsEditable = false,
                AspectGetter = (x) => {
                    bool value = false;
                    var rpd = x as Rpd;
                    if (rpd != null && rpd.EducationalWorks != null) {
                        value = rpd.EducationalWorks.Count == rpd.FormsOfStudy.Count;
                    }
                    return value;
                }
            };
            list.Columns.Add(olvColumnEducationWork);
            //предшествующие дисциплины
            var olvColumnPrevDisciplines = new OLVColumn("Предшествующие дисциплины", nameof(Rpd.PrevDisciplines)) {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnPrevDisciplines);
            //последующие дисциплины
            var olvColumnNextDisciplines = new OLVColumn("Последующие дисциплины", nameof(Rpd.NextDisciplines)) {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnNextDisciplines);
            //ошибки
            var olvColumnErrors = new OLVColumn("Ошибки", nameof(Rpd.Errors)) {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => ((x as Rpd)?.Errors?.IsEmpty ?? true) ? "нет" : "да"
            };
            list.Columns.Add(olvColumnErrors);
            var olvColumnFileName = new OLVColumn("Файл", nameof(Rpd.SourceFileName)) {
                Width = 200,
                IsEditable = false,
                FillsFreeSpace = true
            };
            list.Columns.Add(olvColumnFileName);

            list.PrimarySortColumn = olvColumnDisciplineName;
        }

        /// <summary>
        /// Настройка списка ФОСов
        /// </summary>
        /// <param name="list"></param>
        void TuneFosList(FastObjectListView list) {
            var olvColumnDisciplineName = new OLVColumn("Дисциплина", nameof(Fos.DisciplineName)) {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDisciplineName);

            list.Columns.Add(new OLVColumn("Код", nameof(Fos.DirectionCode)) {
                Width = 80,
                IsEditable = false
            });

            list.Columns.Add(new OLVColumn("Направление подготовки", nameof(Fos.DirectionName)) {
                Width = 200,
                IsEditable = false
            });

            list.Columns.Add(new OLVColumn("Профиль", nameof(Fos.Profile)) {
                Width = 100,
                IsEditable = false
            });

            list.Columns.Add(new OLVColumn("Кафедра", nameof(Fos.Department)) {
                Width = 100,
                IsEditable = false
            });

            list.Columns.Add(new OLVColumn("Год", nameof(Fos.Year)) {
                Width = 50,
                IsEditable = false
            });

            list.Columns.Add(new OLVColumn("Формы обучения", nameof(Fos.FormsOfStudy)) {
                Width = 120,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var fos = x as Fos;
                    if (fos?.FormsOfStudy?.Any() ?? false) {
                        var items = fos.FormsOfStudy.Select(x => x.GetAttribute<EvaluationToolAttribute>()?.ShortDescription ?? x.GetDescription()).ToList();
                        value = string.Join(", ", items);
                    }
                    return value;
                }
            });
            //оценочные средства
            list.Columns.Add(new OLVColumn("Оценочные средства", nameof(Fos.EvalTools)) {
                Width = 120,
                IsEditable = false,
                AspectGetter = x => {
                    var items = (x as Fos)?.EvalTools?
                        .Select(t => t.Key.GetAttribute<EvaluationToolAttribute>()?.ShortDescription ?? t.Key.GetDescription())?
                        .Order()
                        .ToList() ?? ["?"];
                    return string.Join(", ", items);
                }
            });
            //составитель
            list.Columns.Add(new OLVColumn("Составитель", nameof(Fos.Compiler)) {
                Width = 200,
                IsEditable = false
            });
            //матрица компетенций
            list.Columns.Add(new OLVColumn("Компетенции", nameof(Fos.CompetenceMatrix)) {
                Width = 60,
                CheckBoxes = true,
                IsEditable = false,
                AspectGetter = (x) => (x as Fos)?.CompetenceMatrix?.Items?.Any() ?? false
            });
            //паспорт
            list.Columns.Add(new OLVColumn("Паспорт", nameof(Fos.Passport)) {
                Width = 40,
                CheckBoxes = true,
                IsEditable = false,
                AspectGetter = (x) => (x as Fos)?.Passport?.Any() ?? false
            });

            //ошибки
            list.Columns.Add(new OLVColumn("Ошибки", "Errors") {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => ((x as Fos)?.Errors?.Any() ?? false) ? "есть" : "нет"
            });
            //исх. файл
            list.Columns.Add(new OLVColumn("Файл", nameof(Fos.SourceFileName)) {
                Width = 200,
                IsEditable = false,
                FillsFreeSpace = true
            });

            list.PrimarySortColumn = olvColumnDisciplineName;
        }

        void TuneFindAndReplaceList(FastObjectListView list) {
            var olvColFind = new OLVColumn("Найти", "FindPattern") {
                Width = 180,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            list.Columns.Add(olvColFind);
            var olvColReplace = new OLVColumn("Заменить на", "ReplacePattern") {
                Width = 180,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            list.Columns.Add(olvColReplace);
        }

        void TuneDocPropertiesList(FastObjectListView list) {
            var olvColName = new OLVColumn("Свойство", "Name") {
                Width = 150,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            list.Columns.Add(olvColName);
            var olvColValue = new OLVColumn("Значение", "Value") {
                Width = 150,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            list.Columns.Add(olvColValue);
        }

        void TuneFileFixerFileList(FastObjectListView list) {
            var olvColumnDir = new OLVColumn("Директория", "DirectoryName") {
                Width = 500,
                IsEditable = false//, 
                //GroupKeyGetter = x => ((FileFixerItem)x).DirectoryName
            };
            list.Columns.Add(olvColumnDir);
            list.Columns.Add(new OLVColumn("Файл", "FileName") {
                Width = 500,
                //FillsFreeSpace = true,
                IsEditable = false
            });
            list.Columns.Add(new OLVColumn("Тип", "Type") {
                Width = 60,
                //FillsFreeSpace = true,
                IsEditable = false,
                AspectGetter = x => (x as FileFixerItem)?.Type.GetDescription() ?? ""
            });
            //list.AlwaysGroupByColumn = olvColumnDir;
            list.ShowGroups = false;
            list.PrimarySortColumn = olvColumnDir;
            //list.group
        }

        private void FormMain_Load(object sender, EventArgs e) {
            Xceed.Words.NET.Licenser.LicenseKey = "WDN30-W7F00-6RL0S-ERHA";

            App.LoadConfig();

            YaGpt.Init();

            TuneCurriculumList();
            TuneRpdList(fastObjectListViewRpdList);
            TuneFosList(fastObjectListViewFosList);
            TuneDisciplineList(fastObjectListViewDisciplines, true);
            TuneDisciplineList(fastObjectListViewDisciplineListForGeneration, false);
            TuneFindAndReplaceList(fastObjectListViewRpdFixFindAndReplaceItems);
            TuneFindAndReplaceList(fastObjectListViewFosFixFindAndReplace);
            TuneDocPropertiesList(fastObjectListViewRpdFixDocProperties);
            TuneDocPropertiesList(fastObjectListViewFosFixDocProperties);
            TuneFileFixerFileList(fastObjectListViewFileFixerFiles);

            //список шаблонов
            //var templateDir = Path.Combine(Environment.CurrentDirectory, DIR_TEMPLATES);
            if (Templates.Items?.Any() ?? false) {
                var files = Templates.Items.Values;
                var newTemplates = files.Select(f => Path.GetFileName(f)).Except(comboBoxRpdGenTemplates.Items.Cast<string>());
                if (newTemplates.Any()) {
                    comboBoxRpdGenTemplates.Items.AddRange(newTemplates.ToArray());
                }
            }

            //восстановим элементы из конфига
            fastObjectListViewRpdFixFindAndReplaceItems.AddObjects(App.Config.RpdFixFindAndReplaceItems);
            fastObjectListViewRpdFixDocProperties.AddObjects(App.Config.RpdFixDocPropertyList);
            m_matrixFileName = App.Config.CompetenceMatrixFileName ?? "";
            ShowHideFixMode(fastObjectListViewRpdList, splitContainerRpd, false);
            ShowHideFixMode(fastObjectListViewFosList, splitContainerFos, false);
            iconToolStripButtonCurriculaRememberList.Checked = App.Config.StoreCurriculumList;
            iconToolStripButtonCompetenceMatrixAutoload.Checked = App.Config.CompetenceMatrixAutoload;
            iconToolStripButtonRpdRememberList.Checked = App.Config.StoreRpdList;
            checkBoxRpdFixTableOfCompetences.Checked = App.Config.RpdFixTableOfCompetences;
            checkBoxRpdFixSummaryTableOfEduWorks.Checked = App.Config.RpdFixTableOfEduWorks;
            textBoxRpdFixTargetDir.Text = App.Config.RpdFixTargetDir;
            checkBoxRpdFixEduWorkTablesFixTime.Checked = App.Config.RpdFixEduWorkTablesFixTime;
            checkBoxRpdFixEduWorkTablesFixEvalTools.Checked = App.Config.RpdFixEduWorkTablesFixEvalTools;
            checkBoxRpdFixEduWorkTablesFixComptenceResults.Checked = App.Config.RpdFixEduWorkTablesFixCompetenceCodes;
            if (App.Config.RpdFixMaxCompetenceResultsCount >= numericUpDownRpdFixMaxCompetenceResultsCount.Minimum &&
                App.Config.RpdFixMaxCompetenceResultsCount <= numericUpDownRpdFixMaxCompetenceResultsCount.Maximum) {
                numericUpDownRpdFixMaxCompetenceResultsCount.Value = App.Config.RpdFixMaxCompetenceResultsCount;
            }
            textBoxRpdGenFileNameTemplate.Text = App.Config.RpdGenFileNameTemplate;
            textBoxRpdGenTargetDir.Text = App.Config.RpdGenTargetDir;
            textBoxRpdFixFileTemplate.Text = App.Config.RpdFixTemplateFileName;
            checkBoxRpdFixByTemplate.Checked = App.Config.RpdFixByTemplate;
            checkBoxRpdFixSetPrevAndNextDisciplines.Checked = App.Config.RpdFixSetPrevAndNextDisciplines;
            checkBoxRpdFixRemoveColorSelection.Checked = App.Config.RpdFixRemoveColorSelections;
            checkBoxRpdFixFindAndReplace.Checked = App.Config.RpdFixFindAndReplace;
            checkBoxRpdFixEduWorkTablesFullRecreate.Checked = App.Config.RpdFixEduWorkTablesFullRecreate;
            checkBoxRpdFixEduWorksEvalToolsTakeFromFos.Checked = App.Config.RpdFixEduWorkTablesTakeEvalToolsFromFos;

            checkBoxFosFixCompetenceTable1.Checked = App.Config.FosFixCompetenceTable1;
            checkBoxFosFixCompetenceTable2.Checked = App.Config.FosFixCompetenceTable2;
            checkBoxFosFixPassportTable.Checked = App.Config.FosFixPassportTable;
            checkBoxFosFixResetSelection.Checked = App.Config.FosFixResetSelection;
            checkBoxFosFixCompetenceIndicators.Checked = App.Config.FosFixCompetenceIndicators;
            textBoxFosFixTargetDir.Text = App.Config.FosFixTargetDir;
            fastObjectListViewFosFixFindAndReplace.AddObjects(App.Config.FosFixFindAndReplaceItems);
            fastObjectListViewFosFixDocProperties.AddObjects(App.Config.FosFixDocPropertyList);

            toolStripTextBoxFileFixerTargetDir.Text = App.Config.FileFixerLastDirectory;
            toolStripTextBoxFileFixerFind.Text = App.Config.FileFixerFindText;
            checkBoxFileFixerFindAndReplace.Checked = App.Config.FileFixerFindAndReplaceApply;
            checkBoxFileFixerResetSelection.Checked = App.Config.FileFixerResetSelectionApply;
            TuneFindAndReplaceList(fastObjectListViewFileFixerFindAndReplaceItems);
            fastObjectListViewFileFixerFindAndReplaceItems.AddObjects(App.Config.FileFixerFindAndReplaceItems);
            checkBoxFileFixerDocPropsApply.Checked = App.Config.FileFixerDocPropsApply;
            TuneDocPropertiesList(fastObjectListViewFileFixerDocProperties);
            fastObjectListViewFileFixerDocProperties.AddObjects(App.Config.FileFixerDocProps);

            if (!string.IsNullOrEmpty(App.Config.RpdGenTemplate)) {
                if (comboBoxRpdGenTemplates.Items.Contains(App.Config.RpdGenTemplate)) {
                    comboBoxRpdGenTemplates.SelectedItem = App.Config.RpdGenTemplate;
                }
            }
            if (comboBoxRpdGenTemplates.SelectedIndex < 0 && comboBoxRpdGenTemplates.Items.Count > 0) {
                comboBoxRpdGenTemplates.SelectedIndex = 0;
            }

            InitEvalToolsLists();
            UpdateCaption();

            Task.Run(() => {
                this.Invoke(new MethodInvoker(() => {
                    tabControl1.SelectTab(tabPageCompetenceMatrix);
                    if (App.Config.CompetenceMatrixAutoload) {
                        m_matrixFileName = App.Config.CompetenceMatrixFileName;
                        LoadCompetenceMatrix(m_matrixFileName, true);
                    }

                    if (App.Config.CurriculumList?.Any() ?? false) {
                        tabControl1.SelectTab(tabPageСurriculum);
                        LoadCurriculumFilesAsync(App.Config.CurriculumList.ToArray());
                    }
                    if (App.Config.RpdList?.Any() ?? false) {
                        tabControl1.SelectTab(tabPageRpd);
                        LoadRpdFilesAsync(App.Config.RpdList.ToArray());
                    }
                    if (App.Config.FosList?.Any() ?? false) {
                        tabControl1.SelectTab(tabPageFos);
                        LoadFosFilesAsync(App.Config.FosList.ToArray());
                    }
                }));
            });
        }

        /// <summary>
        /// Подготовка списков типов оценочных средств для режима исправления РПД
        /// </summary>
        void InitEvalToolsLists() {
            var items = Enum.GetValues(typeof(EEvaluationTool)).Cast<EEvaluationTool>().Order();
            foreach (EEvaluationTool item in items) {
                var idx = checkedListBoxRpdFixEduWorkTablesEvalTools2ndStageItems.Items.Add(item.GetDescription());
                if (App.Config.RpdFixEduWorkTablesEvalTools2ndStageItems == null &&
                    (item == EEvaluationTool.Essay || item == EEvaluationTool.Paper ||
                    item == EEvaluationTool.Presentation || item == EEvaluationTool.ControlWork) ||
                    (App.Config.RpdFixEduWorkTablesEvalTools2ndStageItems?.Contains(item) ?? false)) {
                    checkedListBoxRpdFixEduWorkTablesEvalTools2ndStageItems.SetItemChecked(idx, true);
                }

                idx = checkedListBoxRpdFixEduWorkTablesEvalTools1stStageItems.Items.Add(item.GetDescription());
                if (App.Config.RpdFixEduWorkTablesEvalTools1stStageItems == null &&
                    (item == EEvaluationTool.Survey || item == EEvaluationTool.Testing)
                    || (App.Config.RpdFixEduWorkTablesEvalTools1stStageItems?.Contains(item) ?? false)) {
                    checkedListBoxRpdFixEduWorkTablesEvalTools1stStageItems.SetItemChecked(idx, true);
                }
            }
        }

        void StatusMessage(string msg, bool isError = false) {
            toolStripStatusLabel1.Text = msg;
            toolStripStatusLabel1.ForeColor = isError ? Color.Red : SystemColors.ControlText;
        }

        /// <summary>
        /// Загрузка матрицы компетенций
        /// </summary>
        /// <param name="fileName"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void LoadCompetenceMatrix(string fileName, bool silent) {
            var loadMatrix = true;

            if (string.IsNullOrEmpty(fileName) || !File.Exists(fileName)) {
                if (!silent) {
                    MessageBox.Show("Файл матрицы компетенций не задан.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                loadMatrix = false;
            }

            if (loadMatrix && App.CompetenceMatrix != null && App.CompetenceMatrix.IsLoaded) {
                if (!silent) {
                    if (MessageBox.Show("Матрица уже загружена.\r\nОбновить?", "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation) != DialogResult.Yes) {
                        loadMatrix = false;
                    }
                }
            }

            if (loadMatrix) {
                var matrix = CompetenceMatrix.LoadFromFile(fileName);
                var errLog = new ConcurrentDictionary<string, List<string>>();

                if (matrix.IsLoaded) {
                    App.SetCompetenceMatrix(matrix);
                    if (!matrix.Errors.IsEmpty) {
                        errLog.TryAdd(fileName, matrix.Errors.Items.Select(e => $"{e}").ToList());
                    }
                    //m_matrixFileName = fileName;

                    ShowMatrix(matrix);
                    //MessageBox.Show("Матрица компетенций успешно загружена.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (errLog.Any()) {
                    WriteErrorLog(errLog, "Матрица_компетенций");
                    //var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                    //if (!Directory.Exists(logDir)) {
                    //    Directory.CreateDirectory(logDir);
                    //}
                    //var dt = DateTime.Now;
                    //var reportFile = Path.Combine(logDir, $"компетенции.{dt:yyyy-MM-dd_HH-mm-ss}.log");
                    //var lines = new List<string>() { fileName };
                    //lines.AddRange(matrix.Errors);
                    //File.WriteAllLines(reportFile, lines, Encoding.UTF8);

                    //if (!silent) {
                    //    StatusMessage($"Во время загрузки обнаружены ошибки ({matrix.Errors}). Они сохранены в журнале {reportFile}", true);
                    //    //var errorList = string.Join("\r\n", matrix.Errors);
                    //    //MessageBox.Show($"Во время загрузки обнаружены ошибки:\r\n{errorList}\r\n\r\nОни сохранены в журнале {reportFile}.", "Загрузка матрицы",
                    //      //              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    //}
                }
                else {
                    StatusMessage("Матрица компетенций успешно загружена");
                }
            }
        }

        /// <summary>
        /// Загрузка УП по списку имен файлов
        /// </summary>
        /// <param name="files"></param>
        /// 
        /*
        void LoadCurriculumFiles(string[] files) {
            //labelExcelFileLoading.Visible = true;
            Application.UseWaitCursor = true;
            //labelExcelFileLoading.Text = "";

            //var report = new List<string>();
            var errLog = new ConcurrentDictionary<string, List<string>>();
            var idx = 1;
            foreach (var file in files) {
                StatusMessage($"Загрузка файлов ({idx} из {files.Length})...");
                if (File.Exists(file)) {
                    var curriculum = Curriculum.LoadFromFile(file);
                    App.AddCurriculum(curriculum);

                    fastObjectListViewCurricula.AddObject(curriculum);
                    idx++;
                    Application.DoEvents();

                    //if (report.Count > 0) report.Add(new string('-', 30));
                    //report.Add(file);
                    //report.AddRange(curriculum.Errors);
                    if (curriculum.Errors.Any()) {
                        errLog.TryAdd(file, curriculum.Errors);
                    }
                }
            }
            if (idx > 1) {
                //labelExcelFileLoading.Text += " завершено.";
            }
            if (errLog.Any()) {
                var logFile = WriteErrorLog(errLog, "УП");

                //MessageBox.Show($"Во время загрузки обнаружены ошибки.\r\nОни сохранены в журнале {reportFile}.", "Внимание",
                //              MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else {
                StatusMessage($"Успешно загружено {files.Length} файл(а,ов)");
            }

            Application.UseWaitCursor = false;
        }
        */

        /// <summary>
        /// Загрузка УП по списку имен файлов (асинхронно)
        /// </summary>
        /// <param name="files"></param>
        private async Task LoadCurriculumFilesAsync(string[] files) {
            Application.UseWaitCursor = true;

            var errLog = new ConcurrentDictionary<string, List<string>>();
            var idx = 0;

            StatusMessage("Загружаем файлы...");

            var tasks = new List<Task>();

            foreach (var file in files) {
                tasks.Add(Task.Run(() => {
                    Interlocked.Increment(ref idx);

                    if (File.Exists(file)) {
                        //проверяем: не загружен ли еще файл?
                        var curr = App.Curricula.Values.FirstOrDefault(r => r.SourceFileName.Equals(file));
                        if (curr != null) {
                            Curriculum.LoadFromFile(file, curr);
                        }
                        else {
                            curr = Curriculum.LoadFromFile(file, null);
                            AddCurriculum(curr);
                        }
                        this.Invoke(new MethodInvoker(() => {
                            fastObjectListViewCurricula.BeginUpdate();
                            var selectedObjects = fastObjectListViewCurricula.SelectedObjects;
                            if (fastObjectListViewCurricula.IndexOf(curr) >= 0) {
                                fastObjectListViewCurricula.UpdateObject(curr);
                            }
                            else {
                                fastObjectListViewCurricula.AddObject(curr);
                            }
                            fastObjectListViewCurricula.EnsureModelVisible(curr);
                            selectedObjects.Add(curr);
                            fastObjectListViewCurricula.SelectedObjects = selectedObjects;
                            fastObjectListViewCurricula.EndUpdate();

                            var msg = $"Загрузка файлов ({idx} из {files.Length})...";
                            //labelLoadRpd.Text = msg;
                            StatusMessage(msg);
                            //if (idx == files.Length) {
                            //    labelLoadRpd.Text += " завершено.";
                            //}
                            Application.DoEvents();
                        }));

                        if (curr.Errors.Any()) {
                            errLog.TryAdd(file, curr.Errors);
                        }
                    }
                }));
            }

            await Task.WhenAll(tasks.ToArray());

            //labelLoadRpd.Text += " завершено.";

            if (errLog.Any()) {
                var logFile = WriteErrorLog(errLog, "УП");
            }
            else {
                StatusMessage($"Успешно загружено {files.Length} файл(а,ов)");
            }

            App.SaveStore(Enums.EStoreElements.Curricula);

            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void fastObjectListViewCurricula_CellToolTipShowing(object sender, ToolTipShowingEventArgs e) {
            if (e.Model is Curriculum curriculum) {
                if (e.Column.AspectName.Equals("Errors")) {
                    e.Text = string.Join("\r\n", curriculum.Errors);
                    e.StandardIcon = ToolTipControl.StandardIcons.Error;
                    e.IsBalloon = true;
                }
            }
        }

        private void fastObjectListViewCurricula_FormatRow(object sender, FormatRowEventArgs e) {
            if (e.Model is Curriculum curriculum) {
                if (curriculum.Errors?.Any() ?? false) {
                    e.Item.BackColor = Color.Pink;
                }
            }
        }

        /// <summary>
        /// Показать список дисциплин учебного плана
        /// </summary>
        /// <param name="curriculum"></param>
        void ShowCurriculumDisciplines(Curriculum curriculum) {
            fastObjectListViewDisciplines.BeginUpdate();
            fastObjectListViewDisciplines.ClearObjects();
            fastObjectListViewDisciplines.SetObjects(curriculum.Disciplines.Values);
            fastObjectListViewDisciplines.EndUpdate();
        }

        private void fastObjectListViewCurricula_ItemActivate(object sender, EventArgs e) {
            if (fastObjectListViewCurricula.FocusedObject is Curriculum curriculum) {
                OpenFile(curriculum.SourceFileName);
            }
        }

        private void fastObjectListViewDisciplines_CellToolTipShowing(object sender, ToolTipShowingEventArgs e) {
            if (e.Model is CurriculumDiscipline discipline) {
                if (e.Column.AspectName.Equals("Errors")) {
                    e.Text = string.Join("\r\n", discipline.Errors.Concat(discipline.ExtraErrors));
                    e.StandardIcon = ToolTipControl.StandardIcons.Error;
                    e.IsBalloon = true;
                }
            }
        }

        private void fastObjectListViewDisciplines_FormatRow(object sender, FormatRowEventArgs e) {
            if (e.Model is CurriculumDiscipline discipline) {
                if ((discipline.Errors?.Any() ?? false) || (discipline.ExtraErrors?.Any() ?? false)) {
                    e.Item.BackColor = Color.Pink;
                }
            }
        }

        /// <summary>
        /// Сброс журнала ошибок
        /// </summary>
        /// <param name="errLog"></param>
        /// <param name="logFilePrefix"></param>
        string WriteErrorLog(IDictionary<string, List<string>> errLog, string logFilePrefix) {
            var logFile = "?";
            try {
                var lines = new List<string>();
                foreach (var item in errLog) {
                    if (lines.Count > 0) lines.Add(new string('-', 30));
                    lines.AddRange(item.Value);
                }

                var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                if (!Directory.Exists(logDir)) {
                    Directory.CreateDirectory(logDir);
                }

                var dt = DateTime.Now;
                logFile = Path.Combine(logDir, $"{logFilePrefix}.{dt:yyyy-MM-dd_HH-mm-ss}.log");

                File.WriteAllLines(logFile, lines, Encoding.UTF8);

                var errCount = errLog.Sum(r => r.Value.Count);
                StatusMessage($"Во время загрузки обнаружены ошибки ({errCount} шт.). См. журнал [{logFile}]", true);
            }
            catch (Exception ex) {
            }

            return logFile;
        }

        /// <summary>
        /// Загрузка РПД-файлов
        /// </summary>
        /// <param name="files"></param>
        private async Task LoadRpdFilesAsync(string[] files) {
            Application.UseWaitCursor = true;

            var errLog = new ConcurrentDictionary<string, List<string>>();
            var idx = 0;

            StatusMessage("Загружаем файлы...");

            var tasks = new List<Task>();

            foreach (var file in files) {
                tasks.Add(Task.Run(() => {
                    Interlocked.Increment(ref idx);

                    if (File.Exists(file)) {
                        //проверяем: не загружен ли еще файл?
                        var rpd = App.RpdList.Values.FirstOrDefault(r => r.SourceFileName.Equals(file));
                        if (rpd != null) {
                            Rpd.LoadFromFile(file, rpd);
                        }
                        else {
                            rpd = Rpd.LoadFromFile(file);
                            AddRpd(rpd);
                        }
                        this.Invoke(new MethodInvoker(() => {
                            fastObjectListViewRpdList.BeginUpdate();
                            var selectedObjects = fastObjectListViewRpdList.SelectedObjects;
                            if (fastObjectListViewRpdList.IndexOf(rpd) >= 0) {
                                fastObjectListViewRpdList.UpdateObject(rpd);
                            }
                            else {
                                fastObjectListViewRpdList.AddObject(rpd);
                            }
                            fastObjectListViewRpdList.EnsureModelVisible(rpd);
                            selectedObjects.Add(rpd);
                            fastObjectListViewRpdList.SelectedObjects = selectedObjects;
                            fastObjectListViewRpdList.EndUpdate();

                            this.UpdateRpdFosMatchIndicators();

                            var msg = $"Загрузка файлов ({idx} из {files.Length})...";
                            //labelLoadRpd.Text = msg;
                            StatusMessage(msg);
                            //if (idx == files.Length) {
                            //    labelLoadRpd.Text += " завершено.";
                            //}
                            Application.DoEvents();
                        }));

                        if (!(rpd?.Errors?.IsEmpty ?? true)) {
                            errLog.TryAdd(file, rpd.Errors.Items.Select(e => $"{e}").ToList());
                        }
                    }
                }));
            }

            await Task.WhenAll(tasks.ToArray());

            //labelLoadRpd.Text += " завершено.";

            if (errLog.Any()) {
                var logFile = WriteErrorLog(errLog, "РПД");
            }
            else {
                StatusMessage($"Успешно загружено {files.Length} файл(а,ов)");
            }
            SaveStore(Enums.EStoreElements.Rpd);

            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        void UpdateRpdFosMatchIndicators() {
            var rpdAndFosAreMatched = App.AreRpdAndFosFilesMatchable();
            Color color = rpdAndFosAreMatched switch {
                true => Color.Green,
                false => Color.Red,
                _ => Color.Black
            };
            toolStripLabelRpdFosIndicator.ForeColor = color;
            toolStripLabelFosRpdIndicator.ForeColor = color;
        }

        /// <summary>
        /// Загрузка ФОС-файлов
        /// </summary>
        /// <param name="files"></param>
        private async Task LoadFosFilesAsync(string[] files) {
            Application.UseWaitCursor = true;

            var errLog = new ConcurrentDictionary<string, List<string>>();
            var idx = 0;

            StatusMessage("Загружаем файлы...");

            var tasks = new List<Task>();

            foreach (var file in files) {
                tasks.Add(Task.Run(() => {
                    Interlocked.Increment(ref idx);

                    if (File.Exists(file)) {
                        //проверяем: не загружен ли еще файл?
                        var fos = App.FosList.Values.FirstOrDefault(r => r.SourceFileName.Equals(file));
                        if (fos != null) {
                            Fos.LoadFromFile(file, fos);
                        }
                        else {
                            fos = Fos.LoadFromFile(file);
                            AddFos(fos);
                        }
                        this.Invoke(new MethodInvoker(() => {
                            fastObjectListViewFosList.BeginUpdate();
                            var selectedObjects = fastObjectListViewFosList.SelectedObjects;
                            if (fastObjectListViewFosList.IndexOf(fos) >= 0) {
                                fastObjectListViewFosList.UpdateObject(fos);
                            }
                            else {
                                fastObjectListViewFosList.AddObject(fos);
                            }
                            fastObjectListViewFosList.EnsureModelVisible(fos);
                            selectedObjects.Add(fos);
                            fastObjectListViewFosList.SelectedObjects = selectedObjects;
                            fastObjectListViewFosList.EndUpdate();

                            this.UpdateRpdFosMatchIndicators();

                            var msg = $"Загрузка файлов ({idx} из {files.Length})...";
                            //labelLoadRpd.Text = msg;
                            StatusMessage(msg);
                            //if (idx == files.Length) {
                            //    labelLoadRpd.Text += " завершено.";
                            //}
                            Application.DoEvents();
                        }));

                        if (fos?.Errors?.Any() ?? false) {
                            errLog.TryAdd(file, fos.Errors);
                        }
                    }
                }));
            }

            await Task.WhenAll(tasks.ToArray());

            //labelLoadRpd.Text += " завершено.";

            if (errLog.Any()) {
                var logFile = WriteErrorLog(errLog, "ФОС");
            }
            else {
                StatusMessage($"Успешно загружено {files.Length} файл(а,ов)");
            }
            SaveStore(Enums.EStoreElements.Fos);

            Application.UseWaitCursor = false;
            Application.DoEvents();
        }

        private void fastObjectListViewRpdList_CellToolTipShowing(object sender, ToolTipShowingEventArgs e) {
            if (e.Model is Rpd rpd) {
                if (e.Column.AspectName.Equals("Errors")) {
                    e.Text = string.Join("\r\n", rpd.Errors.Items.Select(e => $"{e}"));
                    e.StandardIcon = ToolTipControl.StandardIcons.Error;
                    e.IsBalloon = true;
                }
            }
        }

        private void fastObjectListViewRpdList_FormatRow(object sender, FormatRowEventArgs e) {
            if (e.Model is Rpd rpd) {
                if ((rpd.Errors?.IsEmpty ?? true) == false) {
                    e.Item.BackColor = Color.Pink;
                }
            }
        }

        private async void buttonRpdCheck_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var targetDir = textBoxRpdGenTargetDir.Text;
                if (string.IsNullOrEmpty(targetDir)) {
                    targetDir = Path.Combine(Environment.CurrentDirectory, $"Исправленные_РПД_{DateTime.Now:yyyy-MM-dd}");
                }
                App.CheckRdp(rpdList, out var report);

                AddReport("Проверка РПД", report);
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, которые требуется исправить.", "Коррекция РПД", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// Добавить вкладку с отчетом
        /// </summary>
        /// <param name="name"></param>
        /// <param name="html"></param>
        async void AddReport(string name, string html) {
            if (!string.IsNullOrEmpty(html)) {
                try {
                    var tabCount = tabControlReports.TabPages.Cast<TabPage>().Select(t => t.Name.StartsWith(name, StringComparison.CurrentCultureIgnoreCase)).Count();
                    name += $" ({tabCount + 1})  [x]";

                    var tabPageNewReport = new TabPage(name);

                    var webView = new WebView2();
                    await webView.EnsureCoreWebView2Async();

                    if (Encoding.UTF8.GetByteCount(html) > 1_000_000_000) {
                        webView.NavigateToString(html);
                    }
                    else {
                        var repDir = Path.Combine(Environment.CurrentDirectory, DIR_REPORTS);
                        var fileName = Path.Combine(repDir, $"отчёт_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.html");
                        File.WriteAllText(fileName, html, Encoding.UTF8);
                        var uri = new Uri(fileName).AbsoluteUri;
                        webView.CoreWebView2.Navigate(uri);
                    }

                    tabPageNewReport.Controls.Add(webView);
                    webView.Dock = DockStyle.Fill;
                    var webViewInterop = new WebViewInterop();
                    webView.CoreWebView2.AddHostObjectToScript("external", webViewInterop);

                    tabPageNewReport.Tag = html;    //запоминаем на случай сохранения
                    tabControlReports.TabPages.Add(tabPageNewReport);
                    tabControlReports.SelectTab(tabPageNewReport);

                    tabControl1.SelectTab(tabPageReports);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Ошибка при отображении html-отчета:\n\n{ex.Message}\n{ex.StackTrace}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
        }

        private void buttonSaveRpdReport_Click(object sender, EventArgs e) {
            var htmlReport = tabControlReports.SelectedTab?.Tag as string;

            if (!string.IsNullOrEmpty(htmlReport)) {
                var repDir = Path.Combine(Environment.CurrentDirectory, DIR_REPORTS);
                if (!Directory.Exists(repDir)) {
                    Directory.CreateDirectory(repDir);
                }
                var dt = DateTime.Now;
                var pos = tabControlReports.SelectedTab.Text.IndexOf('[');
                var repName = tabControlReports.SelectedTab.Text.Substring(0, pos - 1).Trim();
                var reportFile = Path.Combine(repDir, $"Отчёт.{repName}.{dt:yyyy-MM-dd_HH-mm-ss}.html");
                File.WriteAllText(reportFile, htmlReport, Encoding.UTF8);
                MessageBox.Show($"Отчет сохранён в файл\r\n{reportFile}", "Сохранение отчёта",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        void UpdateCaption() {
            var text = $"{this.Tag} - {tabControl1.SelectedTab.Text.Trim()}";
            FastObjectListView list = null;

            if (tabControl1.SelectedTab == tabPageСurriculum) {
                list = fastObjectListViewCurricula;
            }
            if (tabControl1.SelectedTab == tabPageRpd) {
                list = fastObjectListViewRpdList;
            }
            if (tabControl1.SelectedTab == tabPageFos) {
                list = fastObjectListViewFosList;
            }
            if (tabControl1.SelectedTab == tabPageFileFixer) {
                list = fastObjectListViewFileFixerFiles;
            }

            if (list != null) {
                var count = list.GetItemCount();
                text += $" - {count}";
                var selCount = list.SelectedObjects?.Count ?? 0;
                if (selCount > 0) {
                    text += $" ({selCount})";
                }
            }

            this.Text = text;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e) {
            this.UpdateCaption();

            if (tabControl1.SelectedTab == tabPageRpdGeneration) {
                //матрица компетенций
                if (App.CompetenceMatrix == null) {
                    linkLabelRpdGenCompetenceMatrix.LinkColor = Color.Red;
                    linkLabelRpdGenCompetenceMatrix.VisitedLinkColor = Color.Red;
                    linkLabelRpdGenCompetenceMatrix.Text = "не загружена";
                }
                else {
                    linkLabelRpdGenCompetenceMatrix.LinkColor = Color.Green;
                    linkLabelRpdGenCompetenceMatrix.VisitedLinkColor = Color.Green;
                    linkLabelRpdGenCompetenceMatrix.Text = "загружена";
                }
                //группы УП
                var newGroups = App.CurriculumGroups.Values.Except(comboBoxRpdGenCurriculumGroups.Items.Cast<CurriculumGroup>());
                if (newGroups.Any()) {
                    comboBoxRpdGenCurriculumGroups.Items.AddRange(newGroups.ToArray());
                }

                if (string.IsNullOrEmpty(textBoxRpdGenTargetDir.Text)) {
                    textBoxRpdGenTargetDir.Text = Path.Combine(Environment.CurrentDirectory, $"РПД_{DateTime.Now:yyyy-MM-dd}");
                    textBoxRpdGenTargetDir.SelectionStart = textBoxRpdGenTargetDir.Text.Length - 1;
                }
            }
            if (tabControl1.SelectedTab == tabPageRpd) {
                if (string.IsNullOrEmpty(textBoxRpdFixTargetDir.Text)) {
                    textBoxRpdFixTargetDir.Text = Path.Combine(Environment.CurrentDirectory, $"Исправленные_РПД_{DateTime.Now:yyyy-MM-dd}");
                    textBoxRpdFixTargetDir.SelectionStart = textBoxRpdFixTargetDir.Text.Length - 1;
                }
            }
        }

        private void comboBoxSelectCurriculum_SelectedIndexChanged(object sender, EventArgs e) {
            var curriculumGroup = comboBoxRpdGenCurriculumGroups.SelectedItem as CurriculumGroup;
            if (curriculumGroup != null) {
                fastObjectListViewDisciplineListForGeneration.BeginUpdate();
                fastObjectListViewDisciplineListForGeneration.ClearObjects();
                fastObjectListViewDisciplineListForGeneration.SetObjects(curriculumGroup.Disciplines.Values);
                fastObjectListViewDisciplineListForGeneration.EndUpdate();

                textBoxRpdGenDirectionCode.Text = curriculumGroup.DirectionCode;
                textBoxRpdGenDirectionName.Text = curriculumGroup.DirectionName;
                textBoxRpdGenProfile.Text = curriculumGroup.Profile;
                textBoxRpdGenDepartment.Text = curriculumGroup.DepartmentName;
                textBoxRpdGenFormsOfStudy.Text = curriculumGroup.FormsOfStudyList;
                textBoxRpdGenFSES.Text = curriculumGroup.FSES;
            }
        }

        private void fastObjectListViewDisciplineListForGeneration_CellToolTipShowing(object sender, ToolTipShowingEventArgs e) {
            if (e.Model is CurriculumDiscipline discipline) {
                if (e.Column.AspectName.Equals("Errors")) {
                    e.Text = string.Join("\r\n", discipline.Errors.Concat(discipline.ExtraErrors));
                    e.StandardIcon = ToolTipControl.StandardIcons.Error;
                    e.IsBalloon = true;
                }
            }
        }

        private void fastObjectListViewDisciplineListForGeneration_FormatRow(object sender, FormatRowEventArgs e) {
            if (e.Model is CurriculumDiscipline discipline) {
                if ((discipline.Errors?.Any() ?? false) || (discipline.ExtraErrors?.Any() ?? false)) {
                    e.Item.BackColor = Color.Pink;
                }
                if (fastObjectListViewDisciplineListForGeneration.IsChecked(e.Model)) {
                    e.Item.Font = new System.Drawing.Font(fastObjectListViewDisciplineListForGeneration.Font, FontStyle.Bold);
                }
            }
        }

        private void buttonSelectRpdTargetDir_Click(object sender, EventArgs e) {
            var initDir = Directory.Exists(textBoxRpdGenTargetDir.Text) ? textBoxRpdGenTargetDir.Text : Environment.CurrentDirectory;
            folderBrowserDialogSelectDir.InitialDirectory = initDir;

            if (folderBrowserDialogSelectDir.ShowDialog() == DialogResult.OK) {
                textBoxRpdGenTargetDir.Text = folderBrowserDialogSelectDir.SelectedPath;
            }
        }

        private void buttonGenerate_Click(object sender, EventArgs e) {
            var disciplines = fastObjectListViewDisciplineListForGeneration.CheckedObjects?.Cast<CurriculumDiscipline>().ToList();
            if (disciplines?.Any() ?? false) {
                var disciplineList = string.Join("\r\n  * ", disciplines.Select(d => d.Name));

                if (MessageBox.Show($"Вы уверены, что хотите сгенерировать РПД для следующих дисциплин:\r\n  * {disciplineList}\r\n" +
                                    $"в целевой директории:\r\n{textBoxRpdGenTargetDir.Text}\r\n?\r\n" +
                                    $"Ранее созданные файлы в целевой директории будут перезаписаны.\r\n" +
                                    $"Продолжить?",
                                    "Генерация РПД", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                    var curriculumGroup = comboBoxRpdGenCurriculumGroups.SelectedItem as CurriculumGroup;
                    curriculumGroup.DirectionCode = textBoxRpdGenDirectionCode.Text;
                    curriculumGroup.DirectionName = textBoxRpdGenDirectionName.Text;
                    curriculumGroup.DepartmentName = textBoxRpdGenDepartment.Text;
                    curriculumGroup.Profile = textBoxRpdGenProfile.Text;
                    curriculumGroup.FormsOfStudyList = textBoxRpdGenFormsOfStudy.Text;
                    curriculumGroup.FSES = textBoxRpdGenFSES.Text;

                    curriculumGroup.CheckedDisciplines = disciplines;
                    var rpdTemplate = Path.Combine(Environment.CurrentDirectory, DIR_TEMPLATES, comboBoxRpdGenTemplates.SelectedItem.ToString());

                    labelRpdGenStatus.Visible = true;
                    Application.UseWaitCursor = true;
                    Application.DoEvents();
                    labelRpdGenStatus.Text = "";
                    var targetDir = textBoxRpdGenTargetDir.Text;
                    if (string.IsNullOrEmpty(targetDir)) {
                        targetDir = Path.Combine(Environment.CurrentDirectory, $"РПД_{DateTime.Now:yyyy-MM-dd}");
                    }
                    var newRpdFiles = App.GenerateRpdFiles(curriculumGroup, rpdTemplate, targetDir,
                                                           textBoxRpdGenFileNameTemplate.Text,
                                                           (int idx, CurriculumDiscipline discipline) => {
                                                               this.Invoke(new MethodInvoker(() => {
                                                                   labelRpdGenStatus.Text = $"Генерация РПД для дисциплины\r\n[{discipline.Name}]\r\n" +
                                                                                            $"({idx} из {curriculumGroup.CheckedDisciplines.Count})... ";
                                                                   Application.DoEvents();
                                                               }));
                                                           },
                                                           checkBoxApplyLoadedRpd.Checked,
                                                           out var errors);

                    Application.UseWaitCursor = false;
                    labelRpdGenStatus.Text += " завершено.";
                    var msg = "";
                    if (newRpdFiles.Count > 0) {
                        msg = $"Генерация РПД завершена.\r\nУспешно созданные РПД ({newRpdFiles.Count} шт.):\r\n{string.Join("\r\n", newRpdFiles)}";
                    }
                    else {
                        msg = $"Генерация РПД завершена.\r\nСоздание файлов не удалось.";
                    }
                    if (errors.Any()) {
                        msg += $"\r\n\r\nПолученные ошибки ({errors.Count} шт.):\r\n{string.Join("\r\n", errors)}";
                    }
                    MessageBox.Show(msg, "Генерация РПД", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else {
                MessageBox.Show($"Для генерации РПД необходимо предварительно выбрать группу учебных планов и отметить нужные дисциплины в списке.", "Генерация РПД",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void label9_Click(object sender, EventArgs e) {

        }

        private void linkLabelRpdGenCompetenceMatrix_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            tabControl1.SelectTab(tabPageCompetenceMatrix);
        }

        private void linkLabelRpdPage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            tabControl1.SelectTab(tabPageRpd);
        }

        void ShowHideFixMode(FastObjectListView list, SplitContainer splitContainer, bool show) {
            if (!int.TryParse(list.Tag?.ToString(), out var maxDist)) {
                maxDist = splitContainer.SplitterDistance;
                list.Tag = maxDist;
            }

            //var maxDist = 280;
            var delta = 10;

            list.BeginUpdate();

            if (show) {
                while (splitContainer.SplitterDistance < maxDist) {
                    var newDist = Math.Min(splitContainer.SplitterDistance + delta, maxDist);
                    splitContainer.SplitterDistance = newDist;
                    //Thread.Sleep(10);
                    Application.DoEvents();
                }
            }
            else {
                while (splitContainer.SplitterDistance > 0) {
                    var newDist = Math.Max(splitContainer.SplitterDistance - delta, 0);
                    splitContainer.SplitterDistance = newDist;
                    //Thread.Sleep(10);
                    Application.DoEvents();
                }
            }
            list.EndUpdate();
        }

        private void buttonRpdShowFixMode_Click(object sender, EventArgs e) {
            ShowHideFixMode(fastObjectListViewRpdList, splitContainerRpd, splitContainerRpd.SplitterDistance == 0);
        }

        private void buttonAddFindAndReplaceItem_Click(object sender, EventArgs e) {
            var newItem = new FindAndReplaceItem();
            fastObjectListViewRpdFixFindAndReplaceItems.AddObject(newItem);
            fastObjectListViewRpdFixFindAndReplaceItems.FocusedObject = newItem;
            fastObjectListViewRpdFixFindAndReplaceItems.EnsureModelVisible(newItem);
            fastObjectListViewRpdFixFindAndReplaceItems.EditModel(newItem);
            fastObjectListViewRpdFixFindAndReplaceItems.CheckObject(newItem);

            App.Config.RpdFixFindAndReplaceItems.Add(newItem);
        }

        private void fastObjectListViewRpdFixFindAndReplaceItems_CellEditFinished(object sender, CellEditEventArgs e) {
            App.SaveConfig();
        }

        private void buttonRemoveFindAndReplaceItems_Click(object sender, EventArgs e) {
            if (fastObjectListViewRpdFixFindAndReplaceItems.SelectedObjects.Count > 0) {
                var ret = MessageBox.Show($"Вы уверены, что хотите удалить выделенные элементы ({fastObjectListViewRpdFixFindAndReplaceItems.SelectedObjects.Count} шт.)?",
                    "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (ret == DialogResult.Yes) {
                    foreach (var item in fastObjectListViewRpdFixFindAndReplaceItems.SelectedObjects) {
                        App.Config.RpdFixFindAndReplaceItems.Remove(item as FindAndReplaceItem);
                    }
                    App.SaveConfig();

                    fastObjectListViewRpdFixFindAndReplaceItems.RemoveObjects(fastObjectListViewRpdFixFindAndReplaceItems.SelectedObjects);
                }
            }
        }

        private void FormMain_FormClosed(object sender, FormClosedEventArgs e) {
            App.Config.CurriculumList = [];
            if (App.Config.StoreCurriculumList) {
                App.Config.CurriculumList = App.Curricula?.Values.Select(c => c.SourceFileName).ToList() ?? [];
            }
            App.Config.RpdList = [];
            if (App.Config.StoreRpdList) {
                App.Config.RpdList = App.RpdList?.Values.Select(r => r.SourceFileName).ToList() ?? [];
            }
            App.Config.FosList = [];
            if (App.Config.StoreFosList) {
                App.Config.FosList = App.FosList?.Values.Select(r => r.SourceFileName).ToList() ?? [];
            }
            App.SaveConfig();
        }

        private void buttonCurriculumClearList_Click(object sender, EventArgs e) {
            if (MessageBox.Show("Вы уверены, что хотите очистить список загруженных учебных планов?",
                "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                App.Curricula.Clear();
                fastObjectListViewCurricula.ClearObjects();
            }
        }

        private void tabControlReports_DrawItem(object sender, DrawItemEventArgs e) {
            //e.Graphics.DrawString("x", e.Font, Brushes.Black, e.Bounds.Right - 15, e.Bounds.Top + 4);
            //e.Graphics.DrawString(this.tabControl1.TabPages[e.Index].Text, e.Font, Brushes.Black, e.Bounds.Left + 12, e.Bounds.Top + 4);
            //e.DrawFocusRectangle();
        }

        private void tabControlReports_MouseDown(object sender, MouseEventArgs e) {
            Rectangle r = tabControlReports.GetTabRect(tabControlReports.SelectedIndex);
            Rectangle closeButton = new Rectangle(r.Right - 22, r.Top + 3, 15, 9);
            if (closeButton.Contains(e.Location)) {
                //if (MessageBox.Show("Закрыть вкладку?", "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                var currIdx = tabControlReports.SelectedIndex;
                tabControlReports.TabPages.Remove(tabControlReports.SelectedTab);
                if (currIdx >= tabControlReports.TabPages.Count) {
                    currIdx = tabControlReports.TabPages.Count - 1;
                }
                tabControlReports.SelectedIndex = currIdx;
                //}
            }
        }

        private void tabControlReports_MouseMove(object sender, MouseEventArgs e) {
            //NativeMethods.TCHITTESTINFO HTI = new NativeMethods.TCHITTESTINFO(tabControl.PointToClient(Cursor.Position));
            //int tabID = NativeMethods.SendMessage(tabControl.Handle, NativeMethods.TCM_HITTEST, IntPtr.Zero, ref HTI);
            //return tabID == -1 ? null : tabControl.TabPages[tabID];
            if (tabControlReports.TabPages.Count == 0) return;
            Rectangle r = tabControlReports.GetTabRect(tabControlReports.SelectedIndex);
            Rectangle closeButton = new Rectangle(r.Right - 22, r.Top + 3, 15, 9);
            if (closeButton.Contains(e.Location)) {
                tabControlReports.Cursor = Cursors.Hand;
            }
            else {
                tabControlReports.Cursor = Cursors.Default;
            }
        }

        private void checkBoxRpdFixTableOfCompetences_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixTableOfCompetences = checkBoxRpdFixTableOfCompetences.Checked;
            App.SaveConfig();
        }

        private void checkBoxRpdFixTableOfEduWorks_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixTableOfEduWorks = checkBoxRpdFixSummaryTableOfEduWorks.Checked;
            App.SaveConfig();
        }

        private void buttonRpdFix_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var targetDir = textBoxRpdFixTargetDir.Text;
                if (string.IsNullOrEmpty(targetDir)) {
                    targetDir = Path.Combine(Environment.CurrentDirectory, $"Исправленные_РПД_{DateTime.Now:yyyy-MM-dd}");
                }
                App.FixRpdFiles(rpdList, targetDir, out var htmlReport);

                AddReport("Коррекция РПД", htmlReport);
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, которые требуется исправить.", "Коррекция РПД", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void textBoxRpdGenFileNameTemplate_TextChanged(object sender, EventArgs e) {
            App.Config.RpdGenFileNameTemplate = textBoxRpdGenFileNameTemplate.Text;
            App.SaveConfig();
        }

        private void textBoxRpdFixTargetDir_TextChanged(object sender, EventArgs e) {
            App.Config.RpdFixTargetDir = textBoxRpdFixTargetDir.Text;
            App.SaveConfig();
        }

        private void textBoxRpdGenTargetDir_TextChanged(object sender, EventArgs e) {
            App.Config.RpdGenTargetDir = textBoxRpdGenTargetDir.Text;
            App.SaveConfig();
        }

        private void buttonRpdFixSelectTargetDir_Click(object sender, EventArgs e) {
            var initDir = Directory.Exists(textBoxRpdFixTargetDir.Text) ? textBoxRpdFixTargetDir.Text : Environment.CurrentDirectory;
            folderBrowserDialogSelectDir.InitialDirectory = initDir;

            if (folderBrowserDialogSelectDir.ShowDialog() == DialogResult.OK) {
                textBoxRpdFixTargetDir.Text = folderBrowserDialogSelectDir.SelectedPath;
            }
        }

        private void linkLabelRpdFixSelectFilesToFix_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            fastObjectListViewRpdList.DeselectAll();

            var selectRpdList = new List<Rpd>();
            foreach (var obj in fastObjectListViewRpdList.Objects) {
                var rpd = obj as Rpd;
                if (rpd != null) {
                    var eduWorks = App.GetEducationWorks(rpd, out var curricula);
                    if (eduWorks.Any() && (eduWorks.Count == rpd.FormsOfStudy.Count)) {
                        var disc = curricula.FirstOrDefault().Value.FindDiscipline(rpd.DisciplineName);
                        if (disc != null) {
                            var matrixItems = App.CompetenceMatrix.GetItems(disc.CompetenceList);
                            var indicators = App.CompetenceMatrix.GetAllAchievementCodes(matrixItems);
                            if (disc.CompetenceList.Intersect(indicators).Count() == disc.CompetenceList.Count) {
                                selectRpdList.Add(rpd);
                            }
                        }
                        //
                    }
                }
            }
            fastObjectListViewRpdList.SelectObjects(selectRpdList);
        }

        private void buttonRpdListClear_Click(object sender, EventArgs e) {
            if (MessageBox.Show("Вы уверены, что хотите очистить список загруженных РПД?",
                                "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                App.RpdList.Clear();
                fastObjectListViewRpdList.ClearObjects();
            }
        }

        private void checkBoxRpdFixByTemplate_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixByTemplate = checkBoxRpdFixByTemplate.Checked;
            App.SaveConfig();
        }

        private void textBoxRpdFixFileTemplate_TextChanged(object sender, EventArgs e) {
            App.Config.RpdFixTemplateFileName = textBoxRpdFixFileTemplate.Text;
            App.SaveConfig();
        }

        private void comboBoxRpdGenTemplates_SelectedIndexChanged(object sender, EventArgs e) {
            App.Config.RpdGenTemplate = comboBoxRpdGenTemplates.SelectedItem?.ToString();
            App.SaveConfig();
        }

        private void linkLabelSelectDesciplinesWithLoadedRpd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            fastObjectListViewDisciplineListForGeneration.DeselectAll();
            var rpdList = App.RpdList.Values.ToHashSet();

            var selectObjects = new List<CurriculumDiscipline>();
            foreach (var item in fastObjectListViewDisciplineListForGeneration.Objects) {
                var disc = item as CurriculumDiscipline;
                if (disc != null) {
                    var rpd = App.FindRpd(disc);
                    if (rpd != null) {
                        selectObjects.Add(disc);
                        rpdList.Remove(rpd);
                    }
                }
            }
            fastObjectListViewDisciplineListForGeneration.CheckObjects(selectObjects);
            if (selectObjects.Any()) {
                fastObjectListViewDisciplineListForGeneration.EnsureModelVisible(selectObjects[0]);
                checkBoxApplyLoadedRpd.Checked = true;
            }
            if (rpdList.Count > 0) {
                fastObjectListViewRpdList.SelectObjects(rpdList.ToList());
            }
            if (rpdList.Any()) {
                MessageBox.Show($"Для следующих РПД не нашлось подходящих дисциплин:\n\n" +
                                $"{string.Join("\n", rpdList.Select(r => r.DisciplineName))}\n\n" +
                                $"Они были выделены в списке.", "Внимание!",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void fastObjectListViewDisciplineListForGeneration_ItemChecked(object sender, ItemCheckedEventArgs e) {
            groupBoxRpdGenDisciplineList.Text = $"3. Выбор дисциплин [{fastObjectListViewDisciplineListForGeneration.CheckedObjects?.Count ?? 0}]";
        }

        private void fastObjectListViewRpdList_ItemActivate(object sender, EventArgs e) {
            if (fastObjectListViewRpdList.FocusedObject is Rpd rpd) {
                OpenFile(rpd.SourceFileName);
            }
        }

        private async void buttonYaGptSendQuestion_Click(object sender, EventArgs e) {
            var temp = double.Parse(textBoxYaGptTemp.Text);
            var ret = await YaGpt.TextGenerationAsync(textBoxYaGptSystemText.Text, textBoxYaGptUserText.Text, temp);

            textBoxYaGptAnswer.Text = ret.text ?? ret.error?.message;
        }

        private void button1_Click_1(object sender, EventArgs e) {
            textBoxYaGptSystemText.Text = "ты - профессор университета";
            textBoxYaGptUserText.Text = "какая цель и задачи у дисциплины \"Конъюнктура мирового рынка\"? Не используй разметку Markdown!";
        }

        private void checkBoxRpdFixFillEduWorkTables_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixEduWorkTablesFixTime = checkBoxRpdFixEduWorkTablesFixTime.Checked;
            App.SaveConfig();
        }

        private void checkBoxSetPrevAndNextDisciplines_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixSetPrevAndNextDisciplines = checkBoxRpdFixSetPrevAndNextDisciplines.Checked;
            App.SaveConfig();
        }

        async Task GenerateRelatedDisciplines(List<Rpd> rpdList, bool prevDisciplines) {
            var tasks = new List<Task>();
            var swMain = Stopwatch.StartNew();

            int idx = 0;
            //foreach (var rpd in rpdList) {
            //tasks.Add(Task.Run(async () => {
            //Task task = Task.Run(() => { });
            foreach (var rpd in rpdList) {
                //task = task.ContinueWith(async t => {
                idx++;

                //});
                //task = newTask;

                //Debug.WriteLine($"{idx}: start");

                List<string> names;
                if (prevDisciplines) {
                    var discList = App.GetPossiblePrevDisciplines(rpd);
                    names = discList.Select(d => d.Name).ToList();
                }
                else {
                    var discList = App.GetPossibleNextDisciplines(rpd);
                    names = discList.Select(d => d.Name).ToList();

                    var prevNames = rpd.GetPrevDisciplines();
                    if (prevNames?.Any() ?? false) {
                        names = names.Except(prevNames).ToList();
                    }
                }
                //var discList = prevDisciplines ? App.GetPossiblePrevDisciplines(rpd) : App.GetPossibleNextDisciplines(rpd);
                //var names = discList.Select(d => d.Name).ToList();
                var sw = Stopwatch.StartNew();
                //Debug.WriteLine($"{idx}: GetRelatedDisciplinesAsync call...");
                var discNames = await YaGpt.GetRelatedDisciplinesAsync(rpd.DisciplineName, names);
                //Debug.WriteLine($"{idx}: await GetRelatedDisciplinesAsync - {sw.ElapsedMilliseconds} ms.");
                /*
                var task = YaGpt.GetRelatedDisciplines(rpd.DisciplineName, names);
                task.Wait();
                var discNames = task.Result;
                */

                //Debug.WriteLine($"{idx}: setting disc - discNames.Count = {discNames.Count}");

                names = discNames.TakeRandom(3, 7);
                if (prevDisciplines) {
                    rpd.SetPrevDisciplines(names);
                }
                else {
                    rpd.SetNextDisciplines(names);
                }

                //Debug.WriteLine($"{idx}: form.Invoke...");
                this.Invoke(new MethodInvoker(() => {
                    fastObjectListViewRpdList.UpdateObject(rpd);

                    //var msg = $"Загрузка файлов ({idx} из {files.Length})...";
                    //labelLoadRpd.Text = msg;
                    //StatusMessage(msg);
                    //if (idx == files.Length) {
                    //    labelLoadRpd.Text += " завершено.";
                    //}
                    //Debug.WriteLine($"{idx}: inside form.Invoke");
                    StatusMessage($"Обработано РПД ({idx} из {rpdList.Count}) (последний запрос: {sw.Elapsed}, всего: {swMain.Elapsed})...");
                    Application.DoEvents();
                }));

                //Debug.WriteLine($"{idx}: end");
                //});
            }

            //await task;
        }

        private async void linkLabelRpdFixGeneratePrevDisciplines_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var sw = Stopwatch.StartNew();
                StatusMessage("Выполнение запросов по генерации списков дисциплин...");
                /*
                var tasks = new List<Task>();
                //foreach (var rpd in rpdList) {
                //    tasks.Add(App.GeneratePrevDisciplinesForRpd(rpd));
                //}
                //await Task.WhenAll(tasks);
                ////await rpdList.ForEachAsync(async rpd => {
                //Parallel.ForEach(rpdList, async rpd => {
                foreach (var rpd in rpdList) {
                    //i++;

                    var discList = App.GetPossiblePrevDisciplines(rpd);
                    var names = discList.Select(d => d.Name).ToList();
                    var discNames = await YaGpt.GetRelatedDisciplines(rpd.DisciplineName, names);
                    //task.Wait();
                    var prevNames = discNames.TakeRandom(3, 7);
                    rpd.SetPrevDisciplines(prevNames);
                    //StatusMessage($"Выполнение запросов по генерации списков дисциплин (выполнено {i} из {rpdList.Count}) [{sw.Elapsed}]...");
                }
                */
                await GenerateRelatedDisciplines(rpdList, true);
                StatusMessage($"Запросы по генерации списков дисциплин выполнены ({sw.Elapsed}).");
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, к которым требуется применить функцию.", "Генерация списка дисциплин", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void fastObjectListViewCurricula_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Delete) {
                if ((fastObjectListViewCurricula.SelectedObjects?.Count ?? 0) > 0) {
                    foreach (var item in fastObjectListViewCurricula.SelectedObjects) {
                        App.Curricula.TryRemove((item as Curriculum).SourceFileName, out _);
                    }

                    fastObjectListViewCurricula.RemoveObjects(fastObjectListViewCurricula.SelectedObjects);
                }
            }
        }

        private async void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var sw = Stopwatch.StartNew();
                StatusMessage("Выполнение запросов по генерации списков дисциплин...");
                //var i = 0;
                ////var tasks = new Task();
                //foreach (var rpd in rpdList) {
                //    i++;
                //    var discList = App.GetPossibleNextDisciplines(rpd);
                //    var names = discList.Select(d => d.Name).ToList();
                //    var prevNames = rpd.GetPrevDisciplines();
                //    if (prevNames?.Any() ?? false) {
                //        names = names.Except(prevNames).ToList();
                //    }
                //    var discNames = await YaGpt.GetRelatedDisciplines(rpd.DisciplineName, names);
                //    //task.Wait();
                //    var nextNames = discNames.TakeRandom(3, 7);
                //    rpd.SetNextDisciplines(nextNames);
                //    StatusMessage($"Выполнение запросов по генерации списков дисциплин (выполнено {i} из {rpdList.Count}) [{sw.Elapsed}]...");
                //}
                await GenerateRelatedDisciplines(rpdList, false);
                StatusMessage($"Запросы по генерации списков дисциплин выполнены ({sw.Elapsed}).");
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, к которым требуется применить функцию.", "Генерация списка дисциплин", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void checkBoxRpdFixResetColorSelection_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixRemoveColorSelections = checkBoxRpdFixRemoveColorSelection.Checked;
            App.SaveConfig();
        }

        private void iconButtonRpdRefresh_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var files = rpdList.Select(rpd => rpd.SourceFileName);
                LoadRpdFilesAsync(files.ToArray());
            }
        }

        private void iconToolStripButtonRpdOpen_Click(object sender, EventArgs e) {
            openFileDialogSelectRpd.InitialDirectory = App.Config.RpdLastLocation ?? Environment.CurrentDirectory;

            if (openFileDialogSelectRpd.ShowDialog(this) == DialogResult.OK) {
                var files = openFileDialogSelectRpd.FileNames.Where(x => !App.HasRpdFile(x)).ToArray();

                if (openFileDialogSelectRpd.FileNames.Length > 0) {
                    App.Config.RpdLastLocation = Path.GetDirectoryName(openFileDialogSelectRpd.FileNames[0]);
                    App.SaveConfig();
                }

                LoadRpdFilesAsync(files);
            }
        }

        private void iconToolStripButtonRpdReload_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var files = rpdList.Select(rpd => rpd.SourceFileName);
                LoadRpdFilesAsync(files.ToArray());
            }
        }

        private void iconToolStripButtonRpdClear_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                if (MessageBox.Show($"Вы уверены, что хотите удалить из списка выделенные РПД ({rpdList.Count} шт.)?",
                    "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                    App.RemoveRpd(rpdList);
                    this.UpdateRpdFosMatchIndicators();
                    fastObjectListViewRpdList.RemoveObjects(fastObjectListViewRpdList.SelectedObjects);
                }
            }
        }

        private void iconToolStripButtonRpdToDb_Click(object sender, EventArgs e) {
            //var sw = Stopwatch.StartNew();

            //App.AddLoadedRpdToStore();

            //StatusMessage($"В Стор добавлены РПД ({App.RpdList.Count} шт.) ({sw.Elapsed}). Всего в Сторе РПД: {App.Store.RpdDic.Count} шт.");
        }

        private void iconToolStripButtonRpdFromDb_Click(object sender, EventArgs e) {
            if (MessageBox.Show($"Вы уверены, что хотите загрузить в список РПД из Стора ({App.Store.RpdDic.Count} шт.)?", "Загрузка РПД",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                if (App.Store?.RpdDic?.Any() ?? false) {
                    fastObjectListViewRpdList.BeginUpdate();
                    var selectedObjects = fastObjectListViewRpdList.SelectedObjects;

                    foreach (var rpd in App.Store.RpdDic.Values) {
                        App.AddRpd(rpd, false);

                        fastObjectListViewRpdList.AddObject(rpd);
                        fastObjectListViewRpdList.EnsureModelVisible(rpd);
                        selectedObjects.Add(rpd);
                    }
                    fastObjectListViewRpdList.SelectedObjects = selectedObjects;
                    fastObjectListViewRpdList.EndUpdate();
                }
            }
        }

        private void iconToolStripButtonRpdCheck_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var addedCurriculumCount = 0;
                var missedCurriculumCount = 0;
                //дозагрузка УП по-возможности
                foreach (var rpd in rpdList) {
                    var curricula = App.FindCurricula(rpd);
                    if ((curricula?.Count ?? 0) != rpd.FormsOfStudy.Count) {
                        var curriculaFromStore = App.Store.CurriculaDic.Values.Where(c =>
                            c.DirectionCode.Equals(rpd.DirectionCode) &&
                            c.Profile.Equals(rpd.Profile) &&
                            !curricula.Select(x => x.Value.FormOfStudy).Contains(c.FormOfStudy)
                        );
                        if (curriculaFromStore != null) {
                            foreach (var curr in curriculaFromStore) {
                                App.AddCurriculum(curr, false);
                                addedCurriculumCount++;
                            }
                        }
                        else {
                            missedCurriculumCount++;
                        }
                    }
                }
                var ret = DialogResult.Yes;
                if (addedCurriculumCount > 0 || missedCurriculumCount > 0) {
                    var msg = "Для проверки ФОС\n";
                    if (addedCurriculumCount > 0) {
                        msg += $"были дозагружены из Стора:\n   УП - {addedCurriculumCount} шт.\n";
                    }
                    if (missedCurriculumCount > 0) {
                        msg += $"не найдено:\n   УП - {missedCurriculumCount} шт.\n";
                    }
                    msg += "\nВы уверены, что хотите продолжить?";
                    ret = MessageBox.Show(msg, "Проверка РПД", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                }

                if (ret == DialogResult.Yes) {
                    CheckRdp(rpdList, out var report);

                    AddReport("Проверка РПД", report);
                }
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, которые требуется проверить.", "Проверка РПД", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void iconToolStripButtonRpdFixMode_Click(object sender, EventArgs e) {
            ShowHideFixMode(fastObjectListViewRpdList, splitContainerRpd, splitContainerRpd.SplitterDistance == 0);
        }

        private void iconToolStripButtonRpdRememberList_Click(object sender, EventArgs e) {
            App.Config.StoreRpdList = iconToolStripButtonRpdRememberList.Checked;
            App.SaveConfig();
        }

        private void iconToolStripButtonComptenceMatrixReload_Click(object sender, EventArgs e) {
            if (!string.IsNullOrEmpty(m_matrixFileName)) {
                LoadCompetenceMatrix(m_matrixFileName, false);
            }
        }

        private void iconToolStripButtonCompetenceMatrixOpen_Click(object sender, EventArgs e) {
            if (!string.IsNullOrEmpty(m_matrixFileName) && File.Exists(m_matrixFileName)) {
                openFileDialogSelectCompetenceMatrixFile.FileName = m_matrixFileName;
                openFileDialogSelectCompetenceMatrixFile.InitialDirectory = Path.GetDirectoryName(m_matrixFileName);
            }
            else {
                openFileDialogSelectCompetenceMatrixFile.InitialDirectory = Environment.CurrentDirectory;
            }

            if (openFileDialogSelectCompetenceMatrixFile.ShowDialog(this) == DialogResult.OK) {
                m_matrixFileName = openFileDialogSelectCompetenceMatrixFile.FileName;
                App.SaveConfig();
                iconToolStripButtonComptenceMatrixReload.PerformClick();
            }
        }

        private void iconToolStripButtonCompetenceMatrixAutoload_Click(object sender, EventArgs e) {
            App.Config.CompetenceMatrixAutoload = iconToolStripButtonCompetenceMatrixAutoload.Checked;
            App.SaveConfig();
        }

        private void iconToolStripButtonCurriculaOpen_Click(object sender, EventArgs e) {
            openFileDialogSelectCurriculumFiles.InitialDirectory = App.Config.CurriculumLastLocation ?? Environment.CurrentDirectory;

            if (openFileDialogSelectCurriculumFiles.ShowDialog(this) == DialogResult.OK) {
                if (openFileDialogSelectCurriculumFiles.FileNames.Length > 0) {
                    App.Config.CurriculumLastLocation = Path.GetDirectoryName(openFileDialogSelectCurriculumFiles.FileNames[0]);
                    App.SaveConfig();
                }

                var files = openFileDialogSelectCurriculumFiles.FileNames.Where(x => !App.HasCurriculumFile(x)).ToArray();

                LoadCurriculumFilesAsync(files);
            }
        }

        private void iconToolStripButtonCurriculaClear_Click(object sender, EventArgs e) {
            var curriculaList = fastObjectListViewCurricula.SelectedObjects?.Cast<Curriculum>().ToList();
            if (curriculaList.Any()) {
                if (MessageBox.Show($"Вы уверены, что хотите удалить из списка выделенные Учебные планы ({curriculaList.Count} шт.)?",
                "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                    App.RemoveCurricula(curriculaList);
                    fastObjectListViewCurricula.RemoveObjects(fastObjectListViewCurricula.SelectedObjects);
                }
            }
        }

        private void iconToolStripButtonCurriculaRememberList_Click(object sender, EventArgs e) {
            App.Config.StoreCurriculumList = iconToolStripButtonCurriculaRememberList.Checked;
            App.SaveConfig();
        }

        private void iconToolStripButton2_Click(object sender, EventArgs e) {
            var curriculaList = fastObjectListViewCurricula.SelectedObjects?.Cast<Curriculum>().ToList();
            if (curriculaList.Any()) {
                var files = curriculaList.Select(curr => curr.SourceFileName);
                LoadCurriculumFilesAsync(files.ToArray());
            }
        }

        private void iconToolStripButtonFosOpen_Click(object sender, EventArgs e) {
            openFileDialogFosSelect.InitialDirectory = App.Config.FosLastLocation ?? Environment.CurrentDirectory;

            if (openFileDialogFosSelect.ShowDialog(this) == DialogResult.OK) {
                var files = openFileDialogFosSelect.FileNames.Where(x => !App.HasFosFile(x)).ToArray();

                if (openFileDialogFosSelect.FileNames.Length > 0) {
                    App.Config.FosLastLocation = Path.GetDirectoryName(openFileDialogFosSelect.FileNames[0]);
                    App.SaveConfig();
                }

                LoadFosFilesAsync(files);
            }
        }

        private void fastObjectListViewFosList_CellToolTipShowing(object sender, ToolTipShowingEventArgs e) {
            if (e.Model is Fos fos) {
                if (e.Column.AspectName.Equals("Errors")) {
                    e.Text = string.Join("\r\n", fos.Errors);
                    e.StandardIcon = ToolTipControl.StandardIcons.Error;
                    e.IsBalloon = true;
                }
            }
        }

        private void fastObjectListViewFosList_FormatRow(object sender, FormatRowEventArgs e) {
            if (e.Model is Fos fos) {
                if (fos.Errors?.Any() ?? false) {
                    e.Item.BackColor = Color.Pink;
                }
            }
        }

        private void iconToolStripButtonFosReload_Click(object sender, EventArgs e) {
            var fosList = fastObjectListViewFosList.SelectedObjects?.Cast<Fos>().ToList();
            if (fosList.Any()) {
                var files = fosList.Select(fos => fos.SourceFileName);
                LoadFosFilesAsync(files.ToArray());
            }
        }

        private void iconToolStripButtonForRememberList_Click(object sender, EventArgs e) {
            App.Config.StoreFosList = iconToolStripButtonForRememberList.Checked;
            App.SaveConfig();
        }

        private void iconToolStripButtonFosFromDb_Click(object sender, EventArgs e) {
            if (MessageBox.Show($"Вы уверены, что хотите загрузить в список ФОС из Стора ({App.Store.FosDic.Count} шт.)?", "Загрузка ФОС",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                if (App.Store?.FosDic?.Any() ?? false) {
                    fastObjectListViewFosList.BeginUpdate();
                    var selectedObjects = fastObjectListViewFosList.SelectedObjects;

                    foreach (var fos in App.Store.FosDic.Values) {
                        App.AddFos(fos, false);

                        fastObjectListViewFosList.AddObject(fos);
                        fastObjectListViewFosList.EnsureModelVisible(fos);
                        selectedObjects.Add(fos);
                    }
                    fastObjectListViewFosList.SelectedObjects = selectedObjects;
                    fastObjectListViewFosList.EndUpdate();
                }
            }
        }

        private void iconToolStripButtonFosToDb_Click(object sender, EventArgs e) {
            //var sw = Stopwatch.StartNew();

            //App.AddLoadedFosToStore();

            //StatusMessage($"В Стор добавлены ФОС ({App.FosList.Count} шт.) ({sw.Elapsed}). Всего в Сторе ФОС: {App.Store.FosDic.Count} шт.");
        }

        private void iconToolStripButtonFosCheck_Click(object sender, EventArgs e) {
            var fosList = fastObjectListViewFosList.SelectedObjects?.Cast<Fos>().ToList();
            if (fosList.Any()) {
                var addedRpdCount = 0;
                //var missedRpdCount = 0;
                var addedCurriculumCount = 0;
                var missedCurriculumCount = 0;
                var missedRpdDisciplines = new List<string>();
                //дозагрузка РПД и УП по-возможности
                foreach (var fos in fosList) {
                    if (App.FindRpd(fos) == null) {
                        var rpdFromStore = App.Store.RpdDic.Values.FirstOrDefault(r => r.Key.Equals(fos.Key));
                        if (rpdFromStore != null) {
                            App.AddRpd(rpdFromStore, false);
                            addedRpdCount++;
                        }
                        else {
                            //missedRpdCount++;
                            missedRpdDisciplines.Add(fos.DisciplineName);
                        }
                    }
                    //if (App.FindCurricula(fos) == null) {
                    //    var curriculumFromStore = App.Store.CurriculaDic.Values.FirstOrDefault(c => c.DirectionCode.Equals(fos.DirectionCode) && c.Profile.Equals(fos.Profile));
                    //    if (curriculumFromStore != null) {
                    //        App.AddCurriculum(curriculumFromStore, false);
                    //        addedCurriculumCount++;
                    //    }
                    //    else {
                    //        missedCurriculumCount++;
                    //    }
                    //}
                    var curricula = App.FindCurricula(fos);
                    if ((curricula?.Count ?? 0) != fos.FormsOfStudy.Count) {
                        var curriculaFromStore = App.Store.CurriculaDic.Values.Where(c =>
                            c.DirectionCode.Equals(fos.DirectionCode) &&
                            c.Profile.Equals(fos.Profile) &&
                            !curricula.Select(x => x.Value.FormOfStudy).Contains(c.FormOfStudy)
                        );
                        if (curriculaFromStore != null) {
                            foreach (var curr in curriculaFromStore) {
                                App.AddCurriculum(curr, false);
                                addedCurriculumCount++;
                            }
                        }
                        else {
                            missedCurriculumCount++;
                        }
                    }
                }
                var ret = DialogResult.Yes;
                if (addedCurriculumCount > 0 || addedRpdCount > 0 || missedCurriculumCount > 0 || missedRpdDisciplines.Count > 0) {
                    var msg = "Для проверки ФОС\n";
                    if (addedCurriculumCount > 0 || addedRpdCount > 0) {
                        msg += $"были дозагружены из Стора:\n   УП - {addedCurriculumCount} шт.\n   РПД - {addedRpdCount} шт.\n";
                    }
                    if (missedCurriculumCount > 0) {
                        msg += $"не найдено УП - {missedCurriculumCount} шт.\n";
                    }
                    if (missedRpdDisciplines.Count > 0) {
                        msg += $"не найдено РПД ({missedRpdDisciplines.Count} шт.):\n   * {string.Join("   * ", missedRpdDisciplines)}\n";
                    }
                    msg += "\nВы уверены, что хотите продолжить?";
                    ret = MessageBox.Show(msg, "Проверка ФОС", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                }

                if (ret == DialogResult.Yes) {
                    CheckFos(fosList, out var report);

                    AddReport("Проверка ФОС", report);
                }
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, которые требуется проверить.", "Проверка ФОС", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void iconToolStripButtonCurriculaFromDb_Click(object sender, EventArgs e) {
            if (MessageBox.Show($"Вы уверены, что хотите загрузить в список УП из Стора ({App.Store.CurriculaDic.Count} шт.)?", "Загрузка УП",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                if (App.Store?.CurriculaDic?.Any() ?? false) {
                    fastObjectListViewCurricula.BeginUpdate();
                    var selectedObjects = fastObjectListViewCurricula.SelectedObjects;

                    foreach (var curr in App.Store.CurriculaDic.Values) {
                        App.AddCurriculum(curr, false);

                        fastObjectListViewCurricula.AddObject(curr);
                        fastObjectListViewCurricula.EnsureModelVisible(curr);
                        selectedObjects.Add(curr);
                    }
                    fastObjectListViewCurricula.SelectedObjects = selectedObjects;
                    fastObjectListViewCurricula.EndUpdate();
                }
            }
        }

        private void iconToolStripButtonFosClear_Click(object sender, EventArgs e) {
            var fosList = fastObjectListViewFosList.SelectedObjects?.Cast<Fos>().ToList();
            if (fosList.Any()) {
                if (MessageBox.Show($"Вы уверены, что хотите удалить из списка выделенные ФОС ({fosList.Count} шт.)?",
                    "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                    App.RemoveFos(fosList);
                    this.UpdateRpdFosMatchIndicators();
                    fastObjectListViewFosList.RemoveObjects(fastObjectListViewFosList.SelectedObjects);
                }
            }
        }

        private void fastObjectListViewFosList_ItemActivate(object sender, EventArgs e) {
            if (fastObjectListViewFosList.FocusedObject is Fos fos) {
                OpenFile(fos.SourceFileName);
            }
        }

        private void fastObjectListViewCurricula_SelectedIndexChanged(object sender, EventArgs e) {
            var curriculum = fastObjectListViewCurricula.FocusedObject as Curriculum;
            if (curriculum != null) {
                ShowCurriculumDisciplines(curriculum);
            }
        }

        private void checkBoxFosFixCompetenceTable1_CheckedChanged(object sender, EventArgs e) {
            App.Config.FosFixCompetenceTable1 = checkBoxFosFixCompetenceTable1.Checked;
            App.SaveConfig();
        }

        private void checkBoxFosFixCompetenceTable2_CheckedChanged(object sender, EventArgs e) {
            App.Config.FosFixCompetenceTable2 = checkBoxFosFixCompetenceTable2.Checked;
            App.SaveConfig();
        }

        private void checkBoxFosFixPassportTable_CheckedChanged(object sender, EventArgs e) {
            App.Config.FosFixPassportTable = checkBoxFosFixPassportTable.Checked;
            App.SaveConfig();
        }

        private void checkBoxFosFixResetSelection_CheckedChanged(object sender, EventArgs e) {
            App.Config.FosFixResetSelection = checkBoxFosFixResetSelection.Checked;
            App.SaveConfig();
        }

        private void textBoxFosFixTargetDir_TextChanged(object sender, EventArgs e) {
            App.Config.FosFixTargetDir = textBoxFosFixTargetDir.Text;
            App.SaveConfig();
        }

        private void buttonFosFixSelectTargetDir_Click(object sender, EventArgs e) {
            var initDir = Directory.Exists(textBoxFosFixTargetDir.Text) ? textBoxFosFixTargetDir.Text : Environment.CurrentDirectory;
            folderBrowserDialogSelectDir.InitialDirectory = initDir;

            if (folderBrowserDialogSelectDir.ShowDialog() == DialogResult.OK) {
                textBoxFosFixTargetDir.Text = folderBrowserDialogSelectDir.SelectedPath;
            }
        }

        private void buttonFosFixAddFindAndReplaceItem_Click(object sender, EventArgs e) {
            var newItem = new FindAndReplaceItem();
            fastObjectListViewFosFixFindAndReplace.AddObject(newItem);
            fastObjectListViewFosFixFindAndReplace.FocusedObject = newItem;
            fastObjectListViewFosFixFindAndReplace.EnsureModelVisible(newItem);
            fastObjectListViewFosFixFindAndReplace.EditModel(newItem);
            fastObjectListViewFosFixFindAndReplace.CheckObject(newItem);

            App.Config.FosFixFindAndReplaceItems.Add(newItem);
        }

        private void iconToolStripButtonFosFixMode_Click(object sender, EventArgs e) {
            ShowHideFixMode(fastObjectListViewFosList, splitContainerFos, splitContainerFos.SplitterDistance == 0);
        }

        private void buttonFosFixRemoveFindAndReplaceItem_Click(object sender, EventArgs e) {
            if (fastObjectListViewFosFixFindAndReplace.SelectedObjects.Count > 0) {
                var ret = MessageBox.Show($"Вы уверены, что хотите удалить выделенные элементы ({fastObjectListViewFosFixFindAndReplace.SelectedObjects.Count} шт.)?",
                    "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (ret == DialogResult.Yes) {
                    foreach (var item in fastObjectListViewFosFixFindAndReplace.SelectedObjects) {
                        App.Config.FosFixFindAndReplaceItems.Remove(item as FindAndReplaceItem);
                    }
                    App.SaveConfig();

                    fastObjectListViewFosFixFindAndReplace.RemoveObjects(fastObjectListViewFosFixFindAndReplace.SelectedObjects);
                }
            }
        }

        private void buttonFosFixStart_Click(object sender, EventArgs e) {
            var fosList = fastObjectListViewFosList.SelectedObjects?.Cast<Fos>().ToList();
            if (fosList.Any()) {
                var targetDir = textBoxFosFixTargetDir.Text;
                if (string.IsNullOrEmpty(targetDir)) {
                    targetDir = Path.Combine(Environment.CurrentDirectory, $"Исправленные_ФОС_{DateTime.Now:yyyy-MM-dd}");
                }
                App.FixFosFiles(fosList, targetDir, out var htmlReport);

                AddReport("Коррекция ФОС", htmlReport);
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, которые требуется исправить.", "Коррекция ФОС", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void iconToolStripButtonCurriculaCheck_Click(object sender, EventArgs e) {
            var curriculaList = fastObjectListViewCurricula.SelectedObjects?.Cast<Curriculum>().ToList();
            if (curriculaList.Any()) {

                CheckCurricula(curriculaList, out var report);

                AddReport("Проверка УП", report);
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, которые требуется проверить.", "Проверка УП", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void checkBoxRpdFixEduWorkTablesFixEvalTools_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixEduWorkTablesFixEvalTools = checkBoxRpdFixEduWorkTablesFixEvalTools.Checked;
            App.SaveConfig();
        }

        private void numericUpDownRpdFixMaxCompetenceResultsCount_ValueChanged(object sender, EventArgs e) {
            App.Config.RpdFixMaxCompetenceResultsCount = numericUpDownRpdFixMaxCompetenceResultsCount.Value;
            App.SaveConfig();
        }

        private void checkedListBoxRpdFixEduWorkTablesEvalToolsFirstItems_SelectedIndexChanged(object sender, EventArgs e) {
            var evalToolDic = Enum.GetValues(typeof(EEvaluationTool)).Cast<EEvaluationTool>().ToDictionary(x => x.GetDescription(), x => x);
            var checkedTools = checkedListBoxRpdFixEduWorkTablesEvalTools1stStageItems.CheckedItems.Cast<string>().Select(x => evalToolDic[x]).ToList();

            App.Config.RpdFixEduWorkTablesEvalTools1stStageItems = checkedTools;
            App.SaveConfig();
        }

        private void checkedListBoxRpdFixEduWorkTablesEvalToolsAllItems_SelectedIndexChanged(object sender, EventArgs e) {
            var evalToolDic = Enum.GetValues(typeof(EEvaluationTool)).Cast<EEvaluationTool>().ToDictionary(x => x.GetDescription(), x => x);
            var checkedTools = checkedListBoxRpdFixEduWorkTablesEvalTools2ndStageItems.CheckedItems.Cast<string>().Select(x => evalToolDic[x]).ToList();

            App.Config.RpdFixEduWorkTablesEvalTools2ndStageItems = checkedTools;
            App.SaveConfig();
        }

        private void checkBoxRpdFixFindAndReplace_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixFindAndReplace = checkBoxRpdFixFindAndReplace.Checked;
            App.SaveConfig();

            buttonAddFindAndReplaceItem.Enabled = checkBoxRpdFixFindAndReplace.Checked;
            buttonRemoveFindAndReplaceItems.Enabled = checkBoxRpdFixFindAndReplace.Checked;
            fastObjectListViewRpdFixFindAndReplaceItems.Enabled = checkBoxRpdFixFindAndReplace.Checked;
        }

        private void checkBoxRpdFixEduWorkTablesFixComptenceResults_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixEduWorkTablesFixCompetenceCodes = checkBoxRpdFixEduWorkTablesFixComptenceResults.Checked;
            App.SaveConfig();
        }

        /// <summary>
        /// Добавить файлы в режим универсальной файловой коррекции
        /// </summary>
        /// <param name="files"></param>
        void AddFilesToFileFixerMode(IEnumerable<string> files, EFileType fileType = EFileType.Auto) {
            var oldFiles = fastObjectListViewFileFixerFiles.Objects.Cast<FileFixerItem>().Select(f => f.FullFileName).ToHashSet();
            var newFiles = files.Where(f => !oldFiles.Contains(f) && File.Exists(f));

            if (newFiles.Any()) {
                var newItems = newFiles.Select(f => new FileFixerItem(f, fileType)).ToList();
                fastObjectListViewFileFixerFiles.BeginUpdate();
                fastObjectListViewFileFixerFiles.AddObjects(newItems);
                fastObjectListViewFileFixerFiles.SelectedObjects = newItems;
                fastObjectListViewFileFixerFiles.EnsureModelVisible(newItems.FirstOrDefault());
                fastObjectListViewFileFixerFiles.EndUpdate();
            }
        }

        /// <summary>
        /// Добавить РПД файлы в режим универсальной файловой коррекции
        /// </summary>
        /// <param name="files"></param>
        void AddRpdFilesToFileFixerMode(List<Rpd> rpdList) {
            var oldFiles = fastObjectListViewFileFixerFiles.Objects.Cast<FileFixerItem>().Select(f => f.FullFileName).ToHashSet();
            var newFiles = rpdList.Where(f => !oldFiles.Contains(f.SourceFileName) && File.Exists(f.SourceFileName));

            if (newFiles.Any()) {
                var newItems = newFiles.Select(f => new FileFixerItem(f)).ToList();
                fastObjectListViewFileFixerFiles.BeginUpdate();
                fastObjectListViewFileFixerFiles.AddObjects(newItems);
                fastObjectListViewFileFixerFiles.SelectedObjects = newItems;
                fastObjectListViewFileFixerFiles.EnsureModelVisible(newItems.FirstOrDefault());
                fastObjectListViewFileFixerFiles.EndUpdate();
            }
        }

        /// <summary>
        /// Добавить ФОС файлы в режим универсальной файловой коррекции
        /// </summary>
        /// <param name="files"></param>
        void AddFosFilesToFileFixerMode(List<Fos> fosList) {
            var oldFiles = fastObjectListViewFileFixerFiles.Objects.Cast<FileFixerItem>().Select(f => f.FullFileName).ToHashSet();
            var newFiles = fosList.Where(f => !oldFiles.Contains(f.SourceFileName) && File.Exists(f.SourceFileName));

            if (newFiles.Any()) {
                var newItems = newFiles.Select(f => new FileFixerItem(f)).ToList();
                fastObjectListViewFileFixerFiles.BeginUpdate();
                fastObjectListViewFileFixerFiles.AddObjects(newItems);
                fastObjectListViewFileFixerFiles.SelectedObjects = newItems;
                fastObjectListViewFileFixerFiles.EnsureModelVisible(newItems.FirstOrDefault());
                fastObjectListViewFileFixerFiles.EndUpdate();
            }
        }

        private void iconToolStripButtonFileFixerOpen_Click(object sender, EventArgs e) {
            openFileDialogFileFixer.InitialDirectory = App.Config.FileFixerLastDirectory ?? Environment.CurrentDirectory;

            if (openFileDialogFileFixer.ShowDialog(this) == DialogResult.OK) {
                if (openFileDialogFileFixer.FileNames.Length > 0) {
                    App.Config.FileFixerLastDirectory = Path.GetDirectoryName(openFileDialogFileFixer.FileNames[0]);
                    App.SaveConfig();
                }
                var oldFiles = fastObjectListViewFileFixerFiles.Objects.Cast<FileFixerItem>().Select(f => f.FullFileName);
                var files = openFileDialogFileFixer.FileNames.Except(oldFiles);

                if (files.Any()) {
                    AddFilesToFileFixerMode(files);
                }
            }
        }

        private void iconToolStripButtonFileFixerDelete_Click(object sender, EventArgs e) {
            var files = fastObjectListViewFileFixerFiles.SelectedObjects?.Cast<FileFixerItem>()?.ToList();
            if (files?.Any() ?? false) {
                if (MessageBox.Show($"Вы уверены, что хотите удалить из списка выделенные файлы ({files.Count} шт.)?", "Внимание",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                    fastObjectListViewFileFixerFiles.RemoveObjects(fastObjectListViewFileFixerFiles.SelectedObjects);
                }
            }
        }

        private void checkBoxRpdFixEduWorkTablesFullRecreate_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixEduWorkTablesFullRecreate = checkBoxRpdFixEduWorkTablesFullRecreate.Checked;
            App.SaveConfig();

            if (App.Config.RpdFixEduWorkTablesFullRecreate) {
                checkBoxRpdFixEduWorkTablesFixTime.Checked = true;
                checkBoxRpdFixEduWorkTablesFixEvalTools.Checked = true;
                checkBoxRpdFixEduWorkTablesFixComptenceResults.Checked = true;
            }
        }

        private void checkBoxRpdFixEduWorksEvalToolsTakeFromFos_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixEduWorkTablesTakeEvalToolsFromFos = checkBoxRpdFixEduWorksEvalToolsTakeFromFos.Checked;
            App.SaveConfig();
        }

        private void checkBoxFosFixCompetenceIndicators_CheckedChanged(object sender, EventArgs e) {
            App.Config.FosFixCompetenceIndicators = checkBoxFosFixCompetenceIndicators.Checked;
            App.SaveConfig();
        }

        private void iconButtonRpdGenAbstracts_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var addedCurriculumCount = 0;
                var missedCurriculumCount = 0;
                //дозагрузка УП по-возможности
                foreach (var rpd in rpdList) {
                    var curricula = App.FindCurricula(rpd);
                    if ((curricula?.Count ?? 0) != rpd.FormsOfStudy.Count) {
                        var curriculaFromStore = App.Store.CurriculaDic.Values.Where(c =>
                            c.DirectionCode.Equals(rpd.DirectionCode) &&
                            c.Profile.Equals(rpd.Profile) &&
                            !curricula.Select(x => x.Value.FormOfStudy).Contains(c.FormOfStudy)
                        );
                        if (curriculaFromStore != null) {
                            foreach (var curr in curriculaFromStore) {
                                App.AddCurriculum(curr, false);
                                addedCurriculumCount++;
                            }
                        }
                        else {
                            missedCurriculumCount++;
                        }
                    }
                }
                if (missedCurriculumCount > 0) {
                    MessageBox.Show($"Для генерации Аннотация РПД необходимо предварительно загрузить группу учебных планов.", "Генерация Аннотаций к РПД",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else {
                    if (!Templates.Items.TryGetValue(Templates.TEMPLATE_ABSTRACTS_FOR_RPD, out var mainTemplate) ||
                        !Templates.Items.TryGetValue(Templates.TEMPLATE_ABSTRACT_FOR_DISCIPLINE, out var disciplineTemplate)) {
                        MessageBox.Show($"Обнаружены не все требуемые шаблоны из списка:\n" +
                                        $"{Templates.TEMPLATE_ABSTRACTS_FOR_RPD}\n{Templates.TEMPLATE_ABSTRACT_FOR_DISCIPLINE}.\n" +
                                        $"Генерация Аннотаций к РПД невозможна.",
                                        "Генерация Аннотаций к РПД", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                    var curriculumGroup = App.GetCurriculumGroup(rpdList.FirstOrDefault());
                    if (curriculumGroup == null) {
                        MessageBox.Show($"Не удалось найти группу УП [{rpdList.FirstOrDefault().DirectionCode} {rpdList.FirstOrDefault().Profile}].\n" +
                                        $"Генерация Аннотация к РПД невозможна.",
                            "Генерация Аннотаций к РПД", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                    if (MessageBox.Show($"Вы уверены, что хотите сгенерировать Аннотации к РПД?",
                                        "Генерация Аннотаций к РПД", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                        Application.UseWaitCursor = true;
                        Application.DoEvents();
                        labelRpdGenStatus.Text = "";
                        var sw = Stopwatch.StartNew();
                        var targetDir = textBoxRpdFixTargetDir.Text;
                        var fileName = App.GenerateAbstractsForRpd(curriculumGroup, mainTemplate, disciplineTemplate, targetDir,
                                                                   "{DirectionCode}_{DirectionName}_{Profile}_Аннотации к РПД.docx",
                                                                   (int idx, int max, Rpd rpd, CurriculumDiscipline discipline, string text) => {
                                                                       this.Invoke(new MethodInvoker(() => {
                                                                           StatusMessage($"Генерация Аннотаций к РПД: {text} ({idx} из {max})... ");
                                                                           Application.DoEvents();
                                                                       }));
                                                                   },
                                                                   checkBoxRpdGenAbstractsReplaceMissedPropsForEllipsis.Checked ? textBoxRpdGenAbstractsReplacementForMissedProps.Text : null,
                                                                   out var errors,
                                                                   out var htmlReport);

                        Application.UseWaitCursor = false;
                        labelRpdGenStatus.Text += " завершено.";
                        AddReport("Генерация Аннотаций к РПД", htmlReport);
                        /*
                        var msg = "";
                        if (!string.IsNullOrEmpty(fileName)) {
                            msg = $"Генерация Аннотаций к РПД завершена.\n" +
                                  $"Время работы: {sw.Elapsed})\n" +
                                  $"Итоговый файл: {fileName}\n";
                        }
                        else {
                            msg = $"Генерация Аннотаций к РПД не успешна.\r\nФайл не создан.";
                        }
                        if (errors.Any()) {
                            msg += $"\r\n\r\nПолученные ошибки ({errors.Count} шт.):\r\n{string.Join("\r\n", errors)}";
                        }
                        StatusMessage("Генерация Аннотаций к РПД завершена.");
                        MessageBox.Show(msg, "Генерация Аннотаций к РПД", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        */
                    }
                }
            }
        }

        private void iconToolStripButtonRpdFixToFileFixer_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                AddRpdFilesToFileFixerMode(rpdList);
                tabControl1.SelectedTab = tabPageFileFixer;
            }
        }

        private void iconToolStripButtonFosToFileFixer_Click(object sender, EventArgs e) {
            var fosList = fastObjectListViewFosList.SelectedObjects?.Cast<Fos>().ToList();
            if (fosList.Any()) {
                AddFosFilesToFileFixerMode(fosList);
                tabControl1.SelectedTab = tabPageFileFixer;
            }
        }

        private void iconToolStripButtonFileFixerDistribByDirs_Click(object sender, EventArgs e) {
            var files = fastObjectListViewFileFixerFiles.SelectedObjects?.Cast<FileFixerItem>()?.ToList();
            if (files.Count == 0) return;

            var count = 0;
            var sw = Stopwatch.StartNew();
            var errorList = new ErrorList();

            //1. определим корни
            var roots = new HashSet<string>();
            foreach (var file in files) {
                if (file.Type == EFileType.Rpd) {
                    if (file.Rpd != null) {
                        roots.Add($"{file.Rpd.DirectionCode}_{file.Rpd.DirectionName}_{file.Rpd.Profile}");
                    }
                    else {
                        errorList.AddSimple($"{file.FullFileName} - РПД не загружен");
                    }
                }
                if (file.Type == EFileType.Fos) {
                    if (file.Fos != null) {
                        roots.Add($"{file.Fos.DirectionCode}_{file.Fos.DirectionName}_{file.Fos.Profile}");
                    }
                    else {
                        errorList.AddSimple($"{file.FullFileName} - ФОС не загружен");
                    }
                }
            }
            //2. создание директорий
            foreach (var root in roots) {
                var dirRoot = Path.Combine(App.Config.FileFixerLastDirectory, root);
                if (!Directory.Exists(dirRoot)) Directory.CreateDirectory(dirRoot);

                var rpdDirs = new Dictionary<EDisciplineType, string>() {
                    [EDisciplineType.Required] = Path.Combine(dirRoot, "РПД", "1. Обязательная часть"),
                    [EDisciplineType.Variable] = Path.Combine(dirRoot, "РПД", "2. Часть, формируемая уч_обр_отн"),
                    [EDisciplineType.ByChoice] = Path.Combine(dirRoot, "РПД", "3. Дисциплины по выбору"),
                    [EDisciplineType.Optional] = Path.Combine(dirRoot, "РПД", "4. Факультативы"),
                };
                var fosDirs = new Dictionary<EDisciplineType, string>() {
                    [EDisciplineType.Required] = Path.Combine(dirRoot, "ФОС", "1. Обязательная часть"),
                    [EDisciplineType.Variable] = Path.Combine(dirRoot, "ФОС", "2. Часть, формируемая уч_обр_отн"),
                    [EDisciplineType.ByChoice] = Path.Combine(dirRoot, "ФОС", "3. Дисциплины по выбору"),
                    [EDisciplineType.Optional] = Path.Combine(dirRoot, "ФОС", "4. Факультативы"),
                };

                Directory.CreateDirectory(Path.Combine(dirRoot, "РПД"));
                Directory.CreateDirectory(rpdDirs[EDisciplineType.Required]);
                Directory.CreateDirectory(rpdDirs[EDisciplineType.Variable]);
                Directory.CreateDirectory(rpdDirs[EDisciplineType.ByChoice]);
                Directory.CreateDirectory(rpdDirs[EDisciplineType.Optional]);
                Directory.CreateDirectory(Path.Combine(dirRoot, "ФОС"));
                Directory.CreateDirectory(fosDirs[EDisciplineType.Required]);
                Directory.CreateDirectory(fosDirs[EDisciplineType.Variable]);
                Directory.CreateDirectory(fosDirs[EDisciplineType.ByChoice]);
                Directory.CreateDirectory(fosDirs[EDisciplineType.Optional]);
                Directory.CreateDirectory(Path.Combine(dirRoot, "ГИА"));
                Directory.CreateDirectory(Path.Combine(dirRoot, "Практики"));
                Directory.CreateDirectory(Path.Combine(dirRoot, "Методические рекомендации"));
                //var dirRpd = Path.Combine(dirRoot, "РПД");
                //if (!Directory.Exists(dirRpd)) Directory.CreateDirectory(dirRpd);

                //var dirFos = Path.Combine(dirRoot, "ФОС");
                //if (!Directory.Exists(dirFos)) Directory.CreateDirectory(dirFos);

                //var dirGia = Path.Combine(dirRoot, "ГИА");
                //if (!Directory.Exists(dirGia)) Directory.CreateDirectory(dirGia);

                //var dirGia = Path.Combine(dirRoot, "ГИА");
                //if (!Directory.Exists(dirGia)) Directory.CreateDirectory(dirGia);

                foreach (var f in files) {
                    try {
                        if (f.Type == EFileType.Rpd) {
                            var disc = FindDiscipline(f.Rpd);
                            if (disc != null && disc.Type.HasValue) {
                                var newFileName = Path.Combine(rpdDirs[disc.Type.Value], f.FileName);
                                File.Copy(f.FullFileName, newFileName, true);
                                count++;
                            }
                        }
                        if (f.Type == EFileType.Fos) {
                            var disc = FindDiscipline(f.Fos);
                            if (disc != null && disc.Type.HasValue) {
                                var newFileName = Path.Combine(fosDirs[disc.Type.Value], f.FileName);
                                File.Copy(f.FullFileName, newFileName, true);
                                count++;
                            }
                        }
                    }
                    catch (Exception ex) {
                        errorList.AddException(ex);
                    }
                }
            }
            var html = App.GenerateReport("Распределение файлов", $"Обработано файлов: {count}", sw.Elapsed, errorList, "", "");
            AddReport("Распределение файлов", html);
        }

        private void toolStripTextBoxFileFixerTargetDir_Click(object sender, EventArgs e) {

        }

        private void toolStripTextBoxFileFixerTargetDir_TextChanged(object sender, EventArgs e) {
            App.Config.FileFixerLastDirectory = toolStripTextBoxFileFixerTargetDir.Text;
            App.SaveConfig();
        }

        private void iconToolStripButtonFileFixerSelectTargetDir_Click(object sender, EventArgs e) {
            var initDir = Directory.Exists(toolStripTextBoxFileFixerTargetDir.Text) ? toolStripTextBoxFileFixerTargetDir.Text : Environment.CurrentDirectory;
            folderBrowserDialogSelectDir.InitialDirectory = initDir;

            if (folderBrowserDialogSelectDir.ShowDialog() == DialogResult.OK) {
                toolStripTextBoxFileFixerTargetDir.Text = folderBrowserDialogSelectDir.SelectedPath;
            }
        }

        private void iconToolStripButtonRpdAddDir_Click(object sender, EventArgs e) {
            var initDir = Directory.Exists(App.Config.RpdLastAddDir) ? App.Config.RpdLastAddDir : Environment.CurrentDirectory;
            folderBrowserDialogRpdAdd.InitialDirectory = initDir;

            if (folderBrowserDialogRpdAdd.ShowDialog() == DialogResult.OK) {
                App.Config.RpdLastAddDir = folderBrowserDialogRpdAdd.SelectedPath;

                var files = Directory.EnumerateFiles(App.Config.RpdLastAddDir, "*.docx", new EnumerationOptions() {
                    RecurseSubdirectories = true,
                    MaxRecursionDepth = 100
                }).Except(App.RpdList.Select(r => r.Value.SourceFileName)).ToArray();
                if (files.Any()) {
                    LoadRpdFilesAsync(files);
                }
            }
        }

        private void iconToolStripButtonFosAddDir_Click(object sender, EventArgs e) {
            var initDir = Directory.Exists(App.Config.FosLastAddDir) ? App.Config.FosLastAddDir : Environment.CurrentDirectory;
            folderBrowserDialogFosAdd.InitialDirectory = initDir;

            if (folderBrowserDialogFosAdd.ShowDialog() == DialogResult.OK) {
                App.Config.FosLastAddDir = folderBrowserDialogFosAdd.SelectedPath;

                var files = Directory.EnumerateFiles(App.Config.FosLastAddDir, "*.docx", new EnumerationOptions() {
                    RecurseSubdirectories = true,
                    MaxRecursionDepth = 100
                }).Except(App.FosList.Select(r => r.Value.SourceFileName)).ToArray();
                if (files.Any()) {
                    LoadFosFilesAsync(files);
                }
            }
        }

        private void iconMenuItemRpdReportBase_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var html = Reports.CreateRpdReportBase(rpdList);

                AddReport("Описание РПД", html);
            }
        }

        private void iconMenuItemRpdReportFosMatching_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            var fosList = fastObjectListViewFosList.SelectedObjects?.Cast<Fos>().ToList();

            if (rpdList.Any() && fosList.Any()) {
                var html = Reports.CreateRpdReportFosMatching(rpdList, fosList);

                AddReport("Сопоставление РПД с ФОС", html);
            }
        }

        private void iconMenuItemRpdReportCheckByCurricula_Click(object sender, EventArgs e) {
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            var fosList = fastObjectListViewFosList.SelectedObjects?.Cast<Fos>().ToList();

            if (rpdList.Any() && fosList.Any()) {
                var html = Reports.CreateReportRpdAndFosCheckByCurricula(rpdList, fosList);

                AddReport("Проверка РПД по УП", html);
            }
        }

        void StartProcess(string msg) {
            StatusMessage(msg);
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
        }

        void StopProcess(string msg = "") {
            StatusMessage(msg);
            Application.UseWaitCursor = false;
            Cursor.Current = Cursors.Default;
        }

        private void iconToolStripButtonFileFixerFind_Click(object sender, EventArgs e) {
            var files = fastObjectListViewFileFixerFiles.SelectedObjects?.Cast<FileFixerItem>()?.ToList();
            if (files.Count == 0) return;

            if (!string.IsNullOrEmpty(toolStripTextBoxFileFixerFind.Text)) {
                StartProcess("Поиск файлов...");
                var report = App.FindTextInFiles(toolStripTextBoxFileFixerFind.Text, files);

                AddReport("Поиск текста в файлах", report);
                
                StopProcess("Поиск файлов завершен.");
            }
        }

        private void iconToolStripButtonFileFixerRun_Click(object sender, EventArgs e) {
            var files = fastObjectListViewFileFixerFiles.SelectedObjects?.Cast<FileFixerItem>().ToList();
            if (files?.Any() ?? false) {
                var targetDir = toolStripTextBoxFileFixerTargetDir.Text;
                if (string.IsNullOrEmpty(targetDir)) {
                    targetDir = Path.Combine(Environment.CurrentDirectory, $"Скорректированные_файлы_{DateTime.Now:yyyy-MM-dd}");
                }
                App.FixFiles(files, targetDir, out var htmlReport);

                AddReport("Коррекция файлов", htmlReport);
            }
            else {
                MessageBox.Show("Необходимо выделить файлы, которые требуется скорректировать.", "Коррекция файлов", 
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void toolStripTextBoxFileFixerFind_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter) {
                iconToolStripButtonFileFixerFind.PerformClick();
            }
        }

        private void FormMain_KeyDown(object sender, KeyEventArgs e) {
            if (tabControl1.SelectedTab == tabPageFileFixer) {
                if (e.Control && e.KeyCode == Keys.F) {
                    toolStripTextBoxFileFixerFind.SelectAll();
                    toolStripTextBoxFileFixerFind.Focus();
                }
            }
        }

        private void fastObjectListViewRpdList_SelectedIndexChanged(object sender, EventArgs e) {
        }

        private void fastObjectListViewFosList_SelectedIndexChanged(object sender, EventArgs e) {
        }

        private void fastObjectListViewFileFixerFiles_SelectedIndexChanged(object sender, EventArgs e) {
        }

        private void fastObjectListViewFileFixerFiles_SelectionChanged(object sender, EventArgs e) {
            UpdateCaption();
        }

        private void fastObjectListViewFosList_SelectionChanged(object sender, EventArgs e) {
            UpdateCaption();
        }

        private void fastObjectListViewRpdList_SelectionChanged(object sender, EventArgs e) {
            UpdateCaption();
        }

        private void toolStripTextBoxFileFixerFind_Click(object sender, EventArgs e) {
            App.Config.FileFixerFindText = toolStripTextBoxFileFixerFind.Text;
            App.SaveConfig();
        }

        private void checkBoxFileFixerFindAndReplace_CheckedChanged(object sender, EventArgs e) {
            App.Config.FileFixerFindAndReplaceApply = checkBoxFileFixerFindAndReplace.Checked;
            App.SaveConfig();

            buttonFileFixerFindAndReplaceAddItem.Enabled = checkBoxFileFixerFindAndReplace.Checked;
            buttonFileFixerFindAndReplaceRemoveItem.Enabled = checkBoxFileFixerFindAndReplace.Checked;
            fastObjectListViewFileFixerFindAndReplaceItems.Enabled = checkBoxFileFixerFindAndReplace.Checked;
        }

        private void buttonFileFixerFindAndReplaceAddItem_Click(object sender, EventArgs e) {
            if (fastObjectListViewFileFixerFindAndReplaceItems.SelectedObjects.Count > 0) {
                var ret = MessageBox.Show($"Вы уверены, что хотите удалить выделенные элементы ({fastObjectListViewFileFixerFindAndReplaceItems.SelectedObjects.Count} шт.)?",
                    "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (ret == DialogResult.Yes) {
                    foreach (var item in fastObjectListViewFileFixerFindAndReplaceItems.SelectedObjects) {
                        App.Config.FileFixerFindAndReplaceItems.Remove(item as FindAndReplaceItem);
                    }
                    App.SaveConfig();

                    fastObjectListViewFileFixerFindAndReplaceItems.RemoveObjects(fastObjectListViewFileFixerFindAndReplaceItems.SelectedObjects);
                }
            }
        }

        private void buttonFileFixerFindAndReplaceRemoveItem_Click(object sender, EventArgs e) {
            var newItem = new FindAndReplaceItem();
            fastObjectListViewFileFixerFindAndReplaceItems.AddObject(newItem);
            fastObjectListViewFileFixerFindAndReplaceItems.FocusedObject = newItem;
            fastObjectListViewFileFixerFindAndReplaceItems.EnsureModelVisible(newItem);
            fastObjectListViewFileFixerFindAndReplaceItems.EditModel(newItem);
            fastObjectListViewFileFixerFindAndReplaceItems.CheckObject(newItem);

            App.Config.FileFixerFindAndReplaceItems.Add(newItem);
            App.SaveConfig();
        }

        private void checkBoxFileFixerResetSelection_CheckedChanged(object sender, EventArgs e) {
            App.Config.FileFixerResetSelectionApply = checkBoxFileFixerResetSelection.Checked;
            App.SaveConfig();
        }

        private void checkBoxFileFixerDocPropsApply_CheckedChanged(object sender, EventArgs e) {
            App.Config.FileFixerDocPropsApply = checkBoxFileFixerDocPropsApply.Checked;
            App.SaveConfig();

            fastObjectListViewFileFixerDocProperties.Enabled = checkBoxFileFixerDocPropsApply.Checked;
        }

        private void iconToolStripButtonFileFixerAddDir_Click(object sender, EventArgs e) {
            var initDir = Directory.Exists(App.Config.FileFixerLastAddDir) ? App.Config.FileFixerLastAddDir : Environment.CurrentDirectory;
            folderBrowserDialogFileFixerAdd.InitialDirectory = initDir;

            if (folderBrowserDialogFileFixerAdd.ShowDialog() == DialogResult.OK) {
                App.Config.FileFixerLastAddDir = folderBrowserDialogFileFixerAdd.SelectedPath;

                var oldFiles = fastObjectListViewFileFixerFiles.Objects.Cast<FileFixerItem>().Select(f => f.FullFileName);
                var files = Directory.EnumerateFiles(App.Config.FileFixerLastAddDir, "*.docx", new EnumerationOptions() {
                    RecurseSubdirectories = true,
                    MaxRecursionDepth = 100
                }).Except(oldFiles);
                
                if (files.Any()) {
                    AddFilesToFileFixerMode(files);
                }
            }
        }
    }
}
