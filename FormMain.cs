using BrightIdeasSoftware;
using System.Text;

namespace FosMan {
    public partial class FormMain : Form {
        CompetenceMatrix m_competenceMatrix = null;
        string m_matrixFileName = null;

        public FormMain() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            if (!string.IsNullOrEmpty(m_matrixFileName)) {
                openFileDialog1.FileName = textBoxMatrixFileName.Text;
                openFileDialog1.InitialDirectory = Path.GetDirectoryName(textBoxMatrixFileName.Text);
            }
            else {
                openFileDialog1.InitialDirectory = Environment.CurrentDirectory;
            }

            if (openFileDialog1.ShowDialog(this) == DialogResult.OK) {
                textBoxMatrixFileName.Text = openFileDialog1.FileName;
                //m_matrixFileName = openFileDialog1.FileName;
            }
        }

        private void buttonLoadCompetenceMatrix_Click(object sender, EventArgs e) {
            var loadMatrix = true;

            if (string.IsNullOrEmpty(textBoxMatrixFileName.Text) || !File.Exists(textBoxMatrixFileName.Text)) {
                MessageBox.Show("Файл матрицы компетенций не задан.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                loadMatrix = false;
            }

            if (loadMatrix && m_competenceMatrix != null && m_competenceMatrix.IsLoaded) {
                if (MessageBox.Show("Матрица уже загружена.\r\nОбновить?", "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation) != DialogResult.Yes) {
                    loadMatrix = false;
                }
            }

            if (loadMatrix) {
                m_competenceMatrix = CompetenceMatrix.LoadFromFile(textBoxMatrixFileName.Text);

                if (m_competenceMatrix.IsLoaded) {
                    m_matrixFileName = textBoxMatrixFileName.Text;

                    ShowMatrix(m_competenceMatrix);
                    MessageBox.Show("Матрица компетенций успешно загружена.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else {
                    var errorList = string.Join("\r\n", m_competenceMatrix.Errors);
                    MessageBox.Show($"Обнаружены ошибки:\r\n{errorList}", "Загрузка матрицы", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
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

        void TuneDisciplineList() {
            var olvColumnIndex = new OLVColumn("Индекс", "Index") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewDisciplines.Columns.Add(olvColumnIndex);
            var olvColumnName = new OLVColumn("Наименование", "Name") {
                Width = 300,
                IsEditable = false
            };
            fastObjectListViewDisciplines.Columns.Add(olvColumnName);
            var olvColumnType = new OLVColumn("Тип", "Type") {
                Width = 100,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var discipline = x as CurriculumDiscipline;
                    if (discipline != null) {
                        if (discipline.Type == EDisciplineType.Required) {
                            value = "обязательная";
                        }
                        else if (discipline.Type == EDisciplineType.ByChoice) {
                            value = "по выбору";
                        }
                        else if (discipline.Type == EDisciplineType.Optional) {
                            value = "факультативная";
                        }
                        else {
                            value = "НЕИЗВЕСТНО";
                        }
                    }
                    return value;
                }
            };
            fastObjectListViewDisciplines.Columns.Add(olvColumnType);
            var olvColumnErrors = new OLVColumn("Ошибки", "Errors") {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var discipline = x as CurriculumDiscipline;
                    if (discipline != null) {
                        if (discipline.Errors?.Any() ?? false) {
                            value = "есть";
                        }
                        else {
                            value = "нет";
                        }
                    }
                    return value;
                }
            };
            fastObjectListViewDisciplines.Columns.Add(olvColumnErrors);
            //
            var olvColumnTotalByPlanHours = new OLVColumn("Итого: По плану", "TotalByPlanHours") {
                Width = 50,
                IsEditable = false
            };
            fastObjectListViewDisciplines.Columns.Add(olvColumnTotalByPlanHours);
            var olvColumnTotalContactWorkHours = new OLVColumn("Итого: Конт. раб.", "TotalContactWorkHours") {
                Width = 50,
                IsEditable = false
            };
            fastObjectListViewDisciplines.Columns.Add(olvColumnTotalContactWorkHours);
            var olvColumnTotalSelfStudyHours = new OLVColumn("Итого: СР", "TotalSelfStudyHours") {
                Width = 50,
                IsEditable = false
            };
            fastObjectListViewDisciplines.Columns.Add(olvColumnTotalSelfStudyHours);
            var olvColumnTotalControlHours = new OLVColumn("Итого: КР", "TotalControlHours") {
                Width = 50,
                IsEditable = false
            };
            fastObjectListViewDisciplines.Columns.Add(olvColumnTotalControlHours);
            var colWidth = 55;
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
                fastObjectListViewDisciplines.Columns.Add(olvColumnSemLec);
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
                fastObjectListViewDisciplines.Columns.Add(olvColumnSemLab);
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
                fastObjectListViewDisciplines.Columns.Add(olvColumnSemPractical);
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
                fastObjectListViewDisciplines.Columns.Add(olvColumnSemSelfStudy);
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
                fastObjectListViewDisciplines.Columns.Add(olvColumnSemControl);
            }
        }

        void TuneRpdList() {
            var olvColumnDisciplineName = new OLVColumn("Дисциплина", "DisciplineName") {
                Width = 200,
                IsEditable = false
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnDisciplineName);
            var olvColumnDirectionCode = new OLVColumn("Код", "DirectionCode") {
                Width = 80,
                IsEditable = false
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnDirectionCode);
            var olvColumnDirectionName = new OLVColumn("Направление подготовки", "DirectionName") {
                Width = 200,
                IsEditable = false
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnDirectionName);
            var olvColumnProfile = new OLVColumn("Профиль", "Profile") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnProfile);
            var olvColumnDepartment = new OLVColumn("Кафедра", "Department") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnDepartment);
            var olvColumnYear = new OLVColumn("Год", "Year") {
                Width = 50,
                IsEditable = false
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnYear);
            //var olvColumnCurriculumAcademicYear = new OLVColumn("Учебный год", "AcademicYear") {
            //    Width = 100,
            //    IsEditable = false
            //};
            //fastObjectListViewRpdList.Columns.Add(olvColumnCurriculumAcademicYear);
            var olvColumnFormsOfStudy = new OLVColumn("Формы обучения", "FormOfStudy") {
                Width = 190,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var rpd = x as Rpd;
                    if (rpd != null) {
                        var items = new List<string>();
                        foreach (var item in rpd.FormsOfStudy) {
                            if (item == EFormOfStudy.FullTime) {
                                items.Add("очная");
                            }
                            else if (item == EFormOfStudy.PartTime) {
                                items.Add("заочная");
                            }
                            else if (item == EFormOfStudy.MixedTime) {
                                items.Add("очно-заочная");
                            }
                            else {
                                items.Add("НЕИЗВЕСТНО");
                            }
                        }
                        value = string.Join(", ", items);
                    }
                    return value;
                }
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnFormsOfStudy);
            //матрица компетенций
            var olvColumnCompetenceMatrix = new OLVColumn("Компетенции", "CompetenceMatrix") {
                Width = 60,
                CheckBoxes = true,
                IsEditable = false,
                AspectGetter = (x) => {
                    bool value = false;
                    var rpd = x as Rpd;
                    if (rpd != null && rpd.CompetenceMatrix != null) {
                        value = true;
                    }
                    return value;
                }
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnCompetenceMatrix);
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
            fastObjectListViewRpdList.Columns.Add(olvColumnEducationWork);
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
            fastObjectListViewRpdList.Columns.Add(olvColumnErrors);
            var olvColumnFileName = new OLVColumn("Файл", "SourceFileName") {
                Width = 200,
                IsEditable = false,
                FillsFreeSpace = true
            };
            fastObjectListViewRpdList.Columns.Add(olvColumnFileName);
        }

        private void FormMain_Load(object sender, EventArgs e) {
            TuneCurriculumList();
            TuneDisciplineList();
            TuneRpdList();
        }

        private void buttonSelectExcelFiles_Click(object sender, EventArgs e) {
            if (openFileDialog2.ShowDialog(this) == DialogResult.OK) {
                var files = openFileDialog2.FileNames.Where(x => !App.HasCurriculumFile(x)).ToArray();

                labelExcelFileLoading.Visible = true;
                Application.UseWaitCursor = true;
                labelExcelFileLoading.Text = "";

                var report = new List<string>();
                var idx = 1;
                foreach (var file in files) {
                    labelExcelFileLoading.Text = $"Загрузка файлов ({idx} из {files.Length})...";
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
                if (idx > 1) {
                    labelExcelFileLoading.Text += " завершено.";
                }
                if (report.Any()) {
                    var logDir = Path.Combine(Environment.CurrentDirectory, "Logs");
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
                    e.Text = string.Join("\r\n", discipline.Errors);
                    e.StandardIcon = ToolTipControl.StandardIcons.Error;
                    e.IsBalloon = true;
                }
            }
        }

        private void fastObjectListViewDisciplines_FormatRow(object sender, FormatRowEventArgs e) {
            if (e.Model is CurriculumDiscipline discipline) {
                if (discipline.Errors?.Any() ?? false) {
                    e.Item.BackColor = Color.Pink;
                }
            }
        }

        private void buttonSelectRpdFiles_Click(object sender, EventArgs e) {
            if (openFileDialog3.ShowDialog(this) == DialogResult.OK) {
                var files = openFileDialog3.FileNames.Where(x => !App.HasRpdFile(x)).ToArray();

                labelLoadRpd.Visible = true;
                Application.UseWaitCursor = true;
                labelLoadRpd.Text = "";

                var report = new List<string>();
                var idx = 1;
                foreach (var file in files) {
                    labelLoadRpd.Text = $"Загрузка файлов ({idx} из {files.Length})...";
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
                if (idx > 1) {
                    labelLoadRpd.Text += " завершено.";
                }
                if (report.Any()) {
                    var logDir = Path.Combine(Environment.CurrentDirectory, "Logs");
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

                tabControl1.SelectedTab = tabPageRpdCheck;
            }
            //foreach (var model in fastObjectListViewRpdList.Objects) {
            //    fastObjectListViewRpdList.UpdateObject(model);
            //}
        }
    }
}
