using BrightIdeasSoftware;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace FosMan {
    public partial class FormMain : Form {
        const string DIR_TEMPLATES = "Templates";
        const string DIR_REPORTS = "Reports";
        const string DIR_LOGS = "Logs";

        string m_matrixFileName = null;
        string m_rpdHtmlReport = null;

        public FormMain() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            if (!string.IsNullOrEmpty(textBoxMatrixFileName.Text) && File.Exists(textBoxMatrixFileName.Text)) {
                openFileDialogSelectCompetenceMatrixFile.FileName = textBoxMatrixFileName.Text;
                openFileDialogSelectCompetenceMatrixFile.InitialDirectory = Path.GetDirectoryName(textBoxMatrixFileName.Text);
            }
            else {
                openFileDialogSelectCompetenceMatrixFile.InitialDirectory = Environment.CurrentDirectory;
            }

            if (openFileDialogSelectCompetenceMatrixFile.ShowDialog(this) == DialogResult.OK) {
                textBoxMatrixFileName.Text = openFileDialogSelectCompetenceMatrixFile.FileName;
                //m_matrixFileName = openFileDialog1.FileName;
            }
        }

        private void buttonLoadCompetenceMatrix_Click(object sender, EventArgs e) {
            LoadCompetenceMatrix(textBoxMatrixFileName.Text, false);
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
                        if (curriculum.FormOfStudy == EFormOfStudy.FullTime) {
                            value = EFormOfStudy.FullTime.GetDescription();
                        }
                        else if (curriculum.FormOfStudy == EFormOfStudy.PartTime) {
                            value = EFormOfStudy.PartTime.GetDescription();
                        }
                        else if (curriculum.FormOfStudy == EFormOfStudy.MixedTime) {
                            value = EFormOfStudy.MixedTime.GetDescription();
                        }
                        else {
                            value = EFormOfStudy.Unknown.GetDescription();
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
            //
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
        }

        void TuneRpdList(FastObjectListView list) {
            var olvColumnDisciplineName = new OLVColumn("Дисциплина", "DisciplineName") {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDisciplineName);
            var olvColumnDirectionCode = new OLVColumn("Код", "DirectionCode") {
                Width = 80,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDirectionCode);
            var olvColumnDirectionName = new OLVColumn("Направление подготовки", "DirectionName") {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDirectionName);
            var olvColumnProfile = new OLVColumn("Профиль", "Profile") {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnProfile);
            var olvColumnDepartment = new OLVColumn("Кафедра", "Department") {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDepartment);
            var olvColumnYear = new OLVColumn("Год", "Year") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnYear);
            var olvColumnFormsOfStudy = new OLVColumn("Формы обучения", "FormOfStudy") {
                Width = 190,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var rpd = x as Rpd;
                    if (rpd != null) {
                        var items = rpd.FormsOfStudy.Select(x => x.GetDescription()).ToList();
                        value = string.Join(", ", items);
                    }
                    return value;
                }
            };
            list.Columns.Add(olvColumnFormsOfStudy);
            //матрица компетенций
            var olvColumnCompetenceMatrix = new OLVColumn("Компетенции", "CompetenceMatrix") {
                Width = 60,
                CheckBoxes = true,
                IsEditable = false,
                AspectGetter = (x) => {
                    var rpd = x as Rpd;
                    bool value = rpd?.CompetenceMatrix?.Items?.Any() ?? false;
                    return value;
                }
            };
            list.Columns.Add(olvColumnCompetenceMatrix);
            //объем дисциплины (часы по формам обучения)
            var olvColumnEducationWork = new OLVColumn("Объем", "EducationalWorks") {
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
            //ошибки
            var olvColumnErrors = new OLVColumn("Ошибки", "Errors") {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var rpd = x as Rpd;
                    if (rpd != null) {
                        if (rpd.Errors?.Any() ?? false) {
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
            var olvColumnFileName = new OLVColumn("Файл", "SourceFileName") {
                Width = 200,
                IsEditable = false,
                FillsFreeSpace = true
            };
            list.Columns.Add(olvColumnFileName);
        }

        void TuneRpdFindAndReplaceList() {
            var olvColFind = new OLVColumn("Найти", "FindPattern") {
                Width = 200,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            fastObjectListViewRpdFixFindAndReplaceItems.Columns.Add(olvColFind);
            var olvColReplace = new OLVColumn("Заменить на", "ReplacePattern") {
                Width = 200,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            fastObjectListViewRpdFixFindAndReplaceItems.Columns.Add(olvColReplace);
        }

        private void FormMain_Load(object sender, EventArgs e) {
            App.LoadConfig();

            TuneCurriculumList();
            TuneRpdList(fastObjectListViewRpdList);
            TuneDisciplineList(fastObjectListViewDisciplines, true);
            TuneDisciplineList(fastObjectListViewDisciplineListForGeneration, false);
            TuneRpdFindAndReplaceList();

            //восстановим элементы из конфига
            fastObjectListViewRpdFixFindAndReplaceItems.AddObjects(App.Config.RpdFindAndReplaceItems);
            textBoxMatrixFileName.Text = App.Config.CompetenceMatrixFileName ?? "";
            ShowHideRpdFixMode(false);
            checkBoxStoreCurriculumList.Checked = App.Config.StoreCurriculumList;
            checkBoxCompetenceMatrixAutoload.Checked = App.Config.CompetenceMatrixAutoload;
            checkBoxStoreRpdList.Checked = App.Config.StoreRpdList;

            //отладка
            //textBoxMatrixFileName.Text = @"c:\FosMan\Матрицы_компетенций\test5.docx";
            //buttonLoadCompetenceMatrix.PerformClick();
            //var files = new string[] {
            //    @"c:\FosMan\Учебные_планы\Бизнес- информатика 38.03.05 очное 23.plx.xlsx",
            //    @"c:\FosMan\Учебные_планы\Бизнес-информатика 38.03.05 ОЗ 23.plx.xlsx",
            //    @"c:\FosMan\Учебные_планы\Бизнес-информатика заочн 38.03.05.plx.xlsx"
            //};
            Task.Run(() => {
                this.Invoke(new MethodInvoker(() => {
                    tabControl1.SelectTab(tabPageCompetenceMatrix);
                    if (App.Config.CompetenceMatrixAutoload) {
                        textBoxMatrixFileName.Text = App.Config.CompetenceMatrixFileName;
                        LoadCompetenceMatrix(textBoxMatrixFileName.Text, true);
                    }

                    if (App.Config.CurriculumList?.Any() ?? false) {
                        tabControl1.SelectTab(tabPageСurriculum);
                        LoadCurriculumFiles(App.Config.CurriculumList.ToArray());
                    }
                    if (App.Config.RpdList?.Any() ?? false) {
                        tabControl1.SelectTab(tabPageRpd);
                        LoadRpdFiles(App.Config.RpdList.ToArray());
                    }
                }));
            });
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

                if (matrix.IsLoaded) {
                    App.SetCompetenceMatrix(matrix);
                    m_matrixFileName = fileName;

                    ShowMatrix(matrix);
                    //MessageBox.Show("Матрица компетенций успешно загружена.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (matrix.Errors.Any()) {
                    var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                    if (!Directory.Exists(logDir)) {
                        Directory.CreateDirectory(logDir);
                    }
                    var dt = DateTime.Now;
                    var reportFile = Path.Combine(logDir, $"компетенции.{dt:yyyy-MM-dd_HH-mm-ss}.log");
                    var lines = new List<string>() { fileName };
                    lines.AddRange(matrix.Errors);
                    File.WriteAllLines(reportFile, lines, Encoding.UTF8);

                    if (!silent) {
                        var errorList = string.Join("\r\n", matrix.Errors);
                        MessageBox.Show($"Во время загрузки обнаружены ошибки:\r\n{errorList}\r\n\r\nОни сохранены в журнале {reportFile}.", "Загрузка матрицы",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        /// <summary>
        /// Загрузка УП по списку имен файлов
        /// </summary>
        /// <param name="files"></param>
        void LoadCurriculumFiles(string[] files) {
            labelExcelFileLoading.Visible = true;
            Application.UseWaitCursor = true;
            labelExcelFileLoading.Text = "";

            var report = new List<string>();
            var idx = 1;
            foreach (var file in files) {
                labelExcelFileLoading.Text = $"Загрузка файлов ({idx} из {files.Length})...";
                if (File.Exists(file)) {
                    var curriculum = Curriculum.LoadFromFile(file);
                    App.AddCurriculum(curriculum);

                    fastObjectListViewCurricula.AddObject(curriculum);
                    idx++;
                    Application.DoEvents();

                    if (curriculum.Errors.Any()) {
                        if (report.Count > 0) report.Add(new string('-', 30));
                        report.Add(file);
                        report.AddRange(curriculum.Errors);
                    }
                }
            }
            if (idx > 1) {
                labelExcelFileLoading.Text += " завершено.";
            }
            if (report.Any()) {
                var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                if (!Directory.Exists(logDir)) {
                    Directory.CreateDirectory(logDir);
                }
                var dt = DateTime.Now;
                var reportFile = Path.Combine(logDir, $"уп.{dt:yyyy-MM-dd_HH-mm-ss}.log");
                File.WriteAllLines(reportFile, report, Encoding.UTF8);
                MessageBox.Show($"Во время загрузки обнаружены ошибки.\r\nОни сохранены в журнале {reportFile}.", "Внимание",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            Application.UseWaitCursor = false;
        }

        private void buttonSelectExcelFiles_Click(object sender, EventArgs e) {
            openFileDialogSelectCurriculumFiles.InitialDirectory = App.Config.CurriculumLastLocation ?? Environment.CurrentDirectory;

            if (openFileDialogSelectCurriculumFiles.ShowDialog(this) == DialogResult.OK) {
                if (openFileDialogSelectCurriculumFiles.FileNames.Length > 0) {
                    App.Config.CurriculumLastLocation = Path.GetDirectoryName(openFileDialogSelectCurriculumFiles.FileNames[0]);
                    App.SaveConfig();
                }

                var files = openFileDialogSelectCurriculumFiles.FileNames.Where(x => !App.HasCurriculumFile(x)).ToArray();

                LoadCurriculumFiles(files);
            }
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
            var curriculum = fastObjectListViewCurricula.FocusedObject as Curriculum;
            if (curriculum != null) {
                ShowCurriculumDisciplines(curriculum);
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

        private void buttonSelectRpdFiles_Click(object sender, EventArgs e) {
            openFileDialogSelectRpd.InitialDirectory = App.Config.RpdLastLocation ?? Environment.CurrentDirectory;

            if (openFileDialogSelectRpd.ShowDialog(this) == DialogResult.OK) {
                var files = openFileDialogSelectRpd.FileNames.Where(x => !App.HasRpdFile(x)).ToArray();

                if (openFileDialogSelectRpd.FileNames.Length > 0) {
                    App.Config.RpdLastLocation = Path.GetDirectoryName(openFileDialogSelectRpd.FileNames[0]);
                    App.SaveConfig();
                }

                LoadRpdFiles(files);
            }
        }

        /// <summary>
        /// Загрузка РПД-файлов
        /// </summary>
        /// <param name="files"></param>
        private void LoadRpdFiles(string[] files) {
            labelLoadRpd.Visible = true;
            Application.UseWaitCursor = true;
            labelLoadRpd.Text = "";

            var report = new List<string>();
            var idx = 1;
            foreach (var file in files) {
                labelLoadRpd.Text = $"Загрузка файлов ({idx} из {files.Length})...";
                if (File.Exists(file)) {
                    var rpd = Rpd.LoadFromFile(file);
                    App.AddRpd(rpd);

                    fastObjectListViewRpdList.AddObject(rpd);
                    idx++;
                    Application.DoEvents();

                    if (rpd.Errors.Any()) {
                        if (report.Count > 0) report.Add(new string('-', 30));
                        report.Add(file);
                        report.AddRange(rpd.Errors);
                    }
                }
            }
            if (idx > 1) {
                labelLoadRpd.Text += " завершено.";
            }
            if (report.Any()) {
                var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                if (!Directory.Exists(logDir)) {
                    Directory.CreateDirectory(logDir);
                }
                var dt = DateTime.Now;
                var reportFile = Path.Combine(logDir, $"рпд.{dt:yyyy-MM-dd_HH-mm-ss}.log");
                File.WriteAllLines(reportFile, report, Encoding.UTF8);
                MessageBox.Show($"Во время загрузки обнаружены ошибки.\r\nОни сохранены в журнале {reportFile}.", "Внимание",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            Application.UseWaitCursor = false;
        }

        private void fastObjectListViewRpdList_CellToolTipShowing(object sender, ToolTipShowingEventArgs e) {
            if (e.Model is Rpd rpd) {
                if (e.Column.AspectName.Equals("Errors")) {
                    e.Text = string.Join("\r\n", rpd.Errors);
                    e.StandardIcon = ToolTipControl.StandardIcons.Error;
                    e.IsBalloon = true;
                }
            }
        }

        private void fastObjectListViewRpdList_FormatRow(object sender, FormatRowEventArgs e) {
            if (e.Model is Rpd rpd) {
                if (rpd.Errors?.Any() ?? false) {
                    e.Item.BackColor = Color.Pink;
                }
            }
        }

        private async void buttonRpdCheck_Click(object sender, EventArgs e) {
            App.CheckRdp(out var report);

            if (!string.IsNullOrEmpty(report)) {
                await webView2RpdReport.EnsureCoreWebView2Async();
                webView2RpdReport.NavigateToString(report);

                m_rpdHtmlReport = report;

                tabControl1.SelectedTab = tabPageRpdCheck;
            }
        }

        private void buttonSaveRpdReport_Click(object sender, EventArgs e) {
            if (!string.IsNullOrEmpty(m_rpdHtmlReport)) {
                var repDir = Path.Combine(Environment.CurrentDirectory, DIR_REPORTS);
                if (!Directory.Exists(repDir)) {
                    Directory.CreateDirectory(repDir);
                }
                var dt = DateTime.Now;
                var reportFile = Path.Combine(repDir, $"РПД.{dt:yyyy-MM-dd_HH-mm-ss}.html");
                File.WriteAllText(reportFile, m_rpdHtmlReport, Encoding.UTF8);
                MessageBox.Show($"Отчет сохранён в файл\r\n{reportFile}", "Сохранение отчёта",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e) {
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
                //список шаблонов
                var templateDir = Path.Combine(Environment.CurrentDirectory, DIR_TEMPLATES);
                if (Directory.Exists(templateDir)) {
                    var files = Directory.GetFiles(templateDir, "*.docx");
                    var newTemplates = files.Select(f => Path.GetFileName(f)).Except(comboBoxRpdGenTemplates.Items.Cast<string>());
                    if (newTemplates.Any()) {
                        comboBoxRpdGenTemplates.Items.AddRange(newTemplates.ToArray());
                    }
                    if (comboBoxRpdGenTemplates.SelectedIndex < 0 && comboBoxRpdGenTemplates.Items.Count > 0) {
                        comboBoxRpdGenTemplates.SelectedIndex = 0;
                    }
                }
                if (string.IsNullOrEmpty(textBoxRpdGenTargetDir.Text)) {
                    textBoxRpdGenTargetDir.Text = Path.Combine(Environment.CurrentDirectory, $"РПД_{DateTime.Now:yyyy-MM-dd}");
                    textBoxRpdGenTargetDir.SelectionStart = textBoxRpdGenTargetDir.Text.Length - 1;
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
                textBoxRpdGenDepartment.Text = curriculumGroup.Department;
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
                    e.Item.Font = new Font(fastObjectListViewDisciplineListForGeneration.Font, FontStyle.Bold);
                }
            }
        }

        private void buttonSelectRpdTargetDir_Click(object sender, EventArgs e) {
            var initDir = Directory.Exists(textBoxRpdGenTargetDir.Text) ? textBoxRpdGenTargetDir.Text : Environment.CurrentDirectory;
            folderBrowserDialogRpdTargetDir.InitialDirectory = initDir;

            if (folderBrowserDialogRpdTargetDir.ShowDialog() == DialogResult.OK) {
                textBoxRpdGenTargetDir.Text = folderBrowserDialogRpdTargetDir.SelectedPath;
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
                    curriculumGroup.Department = textBoxRpdGenDepartment.Text;
                    curriculumGroup.Profile = textBoxRpdGenProfile.Text;
                    curriculumGroup.FormsOfStudyList = textBoxRpdGenFormsOfStudy.Text;
                    curriculumGroup.FSES = textBoxRpdGenFSES.Text;

                    curriculumGroup.CheckedDisciplines = disciplines;
                    var rpdTemplate = Path.Combine(Environment.CurrentDirectory, DIR_TEMPLATES, comboBoxRpdGenTemplates.SelectedItem.ToString());

                    labelRpdGenStatus.Visible = true;
                    Application.UseWaitCursor = true;
                    labelRpdGenStatus.Text = "";

                    var newRpdFiles = App.GenerateRpdFiles(curriculumGroup, rpdTemplate, textBoxRpdGenTargetDir.Text,
                                                           textBoxRpdGenFileNameTemplate.Text,
                                                           (int idx, CurriculumDiscipline discipline) => {
                                                               this.Invoke(new MethodInvoker(() => {
                                                                   labelRpdGenStatus.Text = $"Генерация РПД для дисциплины\r\n[{discipline.Name}]\r\n" +
                                                                                            $"({idx} из {curriculumGroup.CheckedDisciplines.Count})... ";
                                                                   Application.DoEvents();
                                                               }));
                                                           },
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

        void ShowHideRpdFixMode(bool show) {
            var maxDist = 200;
            var delta = 10;

            if (show) {
                while (splitContainer1.SplitterDistance < 200) {
                    var newDist = Math.Min(splitContainer1.SplitterDistance + delta, maxDist);
                    splitContainer1.SplitterDistance = newDist;
                    //Thread.Sleep(10);
                    Application.DoEvents();
                }
            }
            else {
                while (splitContainer1.SplitterDistance > 0) {
                    var newDist = Math.Max(splitContainer1.SplitterDistance - delta, 0);
                    splitContainer1.SplitterDistance = newDist;
                    //Thread.Sleep(10);
                    Application.DoEvents();
                }
            }
        }

        private void buttonRpdShowFixMode_Click(object sender, EventArgs e) {
            ShowHideRpdFixMode(splitContainer1.SplitterDistance == 0);
        }

        private void buttonAddFindAndReplaceItem_Click(object sender, EventArgs e) {
            var newItem = new RpdFindAndReplaceItem();
            fastObjectListViewRpdFixFindAndReplaceItems.AddObject(newItem);
            fastObjectListViewRpdFixFindAndReplaceItems.FocusedObject = newItem;
            fastObjectListViewRpdFixFindAndReplaceItems.EnsureModelVisible(newItem);
            fastObjectListViewRpdFixFindAndReplaceItems.EditModel(newItem);
            fastObjectListViewRpdFixFindAndReplaceItems.CheckObject(newItem);

            App.Config.RpdFindAndReplaceItems.Add(newItem);
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
                        App.Config.RpdFindAndReplaceItems.Remove(item as RpdFindAndReplaceItem);
                    }
                    App.SaveConfig();

                    fastObjectListViewRpdFixFindAndReplaceItems.RemoveObjects(fastObjectListViewRpdFixFindAndReplaceItems.SelectedObjects);
                }
            }
        }

        private void textBoxMatrixFileName_TextChanged(object sender, EventArgs e) {
            App.Config.CompetenceMatrixFileName = textBoxMatrixFileName.Text;
            App.SaveConfig();
        }

        private void checkBoxStoreCurriculumList_CheckedChanged(object sender, EventArgs e) {
            App.Config.StoreCurriculumList = checkBoxStoreCurriculumList.Checked;
            App.SaveConfig();
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
            App.SaveConfig();
        }

        private void buttonCurriculumClearList_Click(object sender, EventArgs e) {
            if (MessageBox.Show("Вы уверены, что хотите очистить список загруженных учебных планов?",
                "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                App.Curricula.Clear();
                fastObjectListViewCurricula.ClearObjects();
            }
        }

        private void checkBoxCompetenceMatrixAutoload_CheckedChanged(object sender, EventArgs e) {
            App.Config.CompetenceMatrixAutoload = checkBoxCompetenceMatrixAutoload.Checked;
            App.SaveConfig();
        }

        private void checkBoxStoreRpdList_CheckedChanged(object sender, EventArgs e) {
            App.Config.StoreRpdList = checkBoxStoreRpdList.Checked;
            App.SaveConfig();
        }
    }
}
