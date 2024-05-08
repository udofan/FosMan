using BrightIdeasSoftware;
using System.Text;

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
                MessageBox.Show("���� ������� ����������� �� �����.", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                loadMatrix = false;
            }

            if (loadMatrix && App.CompetenceMatrix != null && App.CompetenceMatrix.IsLoaded) {
                if (MessageBox.Show("������� ��� ���������.\r\n��������?", "��������", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation) != DialogResult.Yes) {
                    loadMatrix = false;
                }
            }

            if (loadMatrix) {
                var matrix = CompetenceMatrix.LoadFromFile(textBoxMatrixFileName.Text);

                if (matrix.IsLoaded) {
                    App.SetCompetenceMatrix(matrix);
                    m_matrixFileName = textBoxMatrixFileName.Text;

                    ShowMatrix(matrix);
                    //MessageBox.Show("������� ����������� ������� ���������.", "��������", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (matrix.Errors.Any()) {
                    var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                    if (!Directory.Exists(logDir)) {
                        Directory.CreateDirectory(logDir);
                    }
                    var dt = DateTime.Now;
                    var reportFile = Path.Combine(logDir, $"�����������.{dt:yyyy-MM-dd_HH-mm-ss}.log");
                    var lines = new List<string>() { textBoxMatrixFileName.Text };
                    lines.AddRange(matrix.Errors);
                    File.WriteAllLines(reportFile, lines, Encoding.UTF8);
                    var errorList = string.Join("\r\n", matrix.Errors);
                    MessageBox.Show($"�� ����� �������� ���������� ������:\r\n{errorList}\r\n\r\n��� ��������� � ������� {reportFile}.", "�������� �������",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        /// <summary>
        /// ������� ������� ����������� � WebView2
        /// </summary>
        /// <param name="errors"></param>
        async void ShowMatrix(CompetenceMatrix matrix) {
            var html = matrix.CreateHtmlReport();

            await webView21.EnsureCoreWebView2Async();
            webView21.NavigateToString(html);
        }

        void TuneCurriculumList() {
            var olvColumnCurriculumIndex = new OLVColumn("���", "DirectionCode") {
                Width = 80,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumIndex);
            var olvColumnCurriculumName = new OLVColumn("����������� ����������", "DirectionName") {
                Width = 200,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumName);
            var olvColumnCurriculumProfile = new OLVColumn("�������", "Profile") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumProfile);
            var olvColumnCurriculumDepartment = new OLVColumn("�������", "Department") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumDepartment);
            var olvColumnCurriculumFaculty = new OLVColumn("���������", "Faculty") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumFaculty);
            var olvColumnCurriculumAcademicYear = new OLVColumn("������� ���", "AcademicYear") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumAcademicYear);
            var olvColumnCurriculumFormOfStudy = new OLVColumn("����� ��������", "FormOfStudy") {
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
            var olvColumnCurriculumErrors = new OLVColumn("������", "Errors") {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var curriculum = x as Curriculum;
                    if (curriculum != null) {
                        if (curriculum.Errors?.Any() ?? false) {
                            value = "����";
                        }
                        else {
                            value = "���";
                        }
                    }
                    return value;
                }
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumErrors);
            var olvColumnCurriculumFileName = new OLVColumn("����", "SourceFileName") {
                Width = 200,
                IsEditable = false,
                FillsFreeSpace = true
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumFileName);
        }

        void TuneDisciplineList(FastObjectListView list, bool addSemesterColumns) {
            var olvColumnIndex = new OLVColumn("������", "Index") {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnIndex);
            var olvColumnName = new OLVColumn("������������", "Name") {
                Width = 300,
                IsEditable = false
            };
            list.Columns.Add(olvColumnName);
            var olvColumnType = new OLVColumn("���", "Type") {
                Width = 100,
                IsEditable = false,
                AspectGetter = (x) => {
                    var discipline = x as CurriculumDiscipline;
                    var value = discipline?.Type?.GetDescription();
                    return value;
                }
            };
            list.Columns.Add(olvColumnType);
            var olvColumnErrors = new OLVColumn("������", "Errors") {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var discipline = x as CurriculumDiscipline;
                    if (discipline != null) {
                        if (discipline.Errors?.Any() ?? false) {
                            value = "����";
                        }
                        else {
                            value = "���";
                        }
                    }
                    return value;
                }
            };
            list.Columns.Add(olvColumnErrors);
            //
            var olvColumnTotalByPlanHours = new OLVColumn("�����: �� �����", "TotalByPlanHours") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnTotalByPlanHours);
            var olvColumnTotalContactWorkHours = new OLVColumn("�����: ����. ���.", "TotalContactWorkHours") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnTotalContactWorkHours);
            var olvColumnTotalSelfStudyHours = new OLVColumn("�����: ��", "TotalSelfStudyHours") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnTotalSelfStudyHours);
            var olvColumnTotalControlHours = new OLVColumn("�����: ��", "TotalControlHours") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnTotalControlHours);
            var colWidth = 55;
            if (addSemesterColumns) {
                for (var i = 0; i < CurriculumDiscipline.SEMESTER_COUNT; i++) {
                    var colIdx = i;
                    var semTitle = $"���. {i + 1}: ";
                    var olvColumnSemTotal = new OLVColumn($"{semTitle}�����", "") {
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
                    var olvColumnSemLec = new OLVColumn($"{semTitle}���", "") {
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
                    var olvColumnSemLab = new OLVColumn($"{semTitle}���", "") {
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
                    var olvColumnSemPractical = new OLVColumn($"{semTitle}��", "") {
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
                    var olvColumnSemSelfStudy = new OLVColumn($"{semTitle}��", "") {
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
                    var olvColumnSemControl = new OLVColumn($"{semTitle}��������", "") {
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
            var olvColumnTotalCompetenceList = new OLVColumn("�����������", "CompetenceList") {
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
            var olvColumnDisciplineName = new OLVColumn("����������", "DisciplineName") {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDisciplineName);
            var olvColumnDirectionCode = new OLVColumn("���", "DirectionCode") {
                Width = 80,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDirectionCode);
            var olvColumnDirectionName = new OLVColumn("����������� ����������", "DirectionName") {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDirectionName);
            var olvColumnProfile = new OLVColumn("�������", "Profile") {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnProfile);
            var olvColumnDepartment = new OLVColumn("�������", "Department") {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDepartment);
            var olvColumnYear = new OLVColumn("���", "Year") {
                Width = 50,
                IsEditable = false
            };
            list.Columns.Add(olvColumnYear);
            var olvColumnFormsOfStudy = new OLVColumn("����� ��������", "FormOfStudy") {
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
            //������� �����������
            var olvColumnCompetenceMatrix = new OLVColumn("�����������", "CompetenceMatrix") {
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
            //����� ���������� (���� �� ������ ��������)
            var olvColumnEducationWork = new OLVColumn("�����", "EducationalWorks") {
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
            //������
            var olvColumnErrors = new OLVColumn("������", "Errors") {
                Width = 50,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var rpd = x as Rpd;
                    if (rpd != null) {
                        if (rpd.Errors?.Any() ?? false) {
                            value = "����";
                        }
                        else {
                            value = "���";
                        }
                    }
                    return value;
                }
            };
            list.Columns.Add(olvColumnErrors);
            var olvColumnFileName = new OLVColumn("����", "SourceFileName") {
                Width = 200,
                IsEditable = false,
                FillsFreeSpace = true
            };
            list.Columns.Add(olvColumnFileName);
        }

        private void FormMain_Load(object sender, EventArgs e) {
            TuneCurriculumList();
            TuneRpdList(fastObjectListViewRpdList);
            TuneDisciplineList(fastObjectListViewDisciplines, true);
            TuneDisciplineList(fastObjectListViewDisciplineListForGeneration, false);
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
                    labelExcelFileLoading.Text = $"�������� ������ ({idx} �� {files.Length})...";
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
                    labelExcelFileLoading.Text += " ���������.";
                }
                if (report.Any()) {
                    var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                    if (!Directory.Exists(logDir)) {
                        Directory.CreateDirectory(logDir);
                    }
                    var dt = DateTime.Now;
                    var reportFile = Path.Combine(logDir, $"��.{dt:yyyy-MM-dd_HH-mm-ss}.log");
                    File.WriteAllLines(reportFile, report, Encoding.UTF8);
                    MessageBox.Show($"�� ����� �������� ���������� ������.\r\n��� ��������� � ������� {reportFile}.", "��������",
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
        /// �������� ������ ��������� �������� �����
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
                    labelLoadRpd.Text = $"�������� ������ ({idx} �� {files.Length})...";
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
                    labelLoadRpd.Text += " ���������.";
                }
                if (report.Any()) {
                    var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                    if (!Directory.Exists(logDir)) {
                        Directory.CreateDirectory(logDir);
                    }
                    var dt = DateTime.Now;
                    var reportFile = Path.Combine(logDir, $"���.{dt:yyyy-MM-dd_HH-mm-ss}.log");
                    File.WriteAllLines(reportFile, report, Encoding.UTF8);
                    MessageBox.Show($"�� ����� �������� ���������� ������.\r\n��� ��������� � ������� {reportFile}.", "��������",
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
                var reportFile = Path.Combine(repDir, $"���.{dt:yyyy-MM-dd_HH-mm-ss}.html");
                File.WriteAllText(reportFile, m_rpdHtmlReport, Encoding.UTF8);
                MessageBox.Show($"����� ������� � ����\r\n{reportFile}", "���������� ������",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e) {
            if (tabControl1.SelectedTab == tabPageRpdGeneration) {
                //������� �����������
                if (App.CompetenceMatrix == null) {
                    linkLabelRpdGenCompetenceMatrix.LinkColor = Color.Red;
                    linkLabelRpdGenCompetenceMatrix.VisitedLinkColor = Color.Red;
                    linkLabelRpdGenCompetenceMatrix.Text = "�� ���������";
                }
                else {
                    linkLabelRpdGenCompetenceMatrix.LinkColor = Color.Green;
                    linkLabelRpdGenCompetenceMatrix.VisitedLinkColor = Color.Green;
                    linkLabelRpdGenCompetenceMatrix.Text = "���������";
                }
                //������ ��
                var newGroups = App.CurriculumGroups.Values.Except(comboBoxRpdGenCurriculumGroups.Items.Cast<CurriculumGroup>());
                if (newGroups.Any()) {
                    comboBoxRpdGenCurriculumGroups.Items.AddRange(newGroups.ToArray());
                }
                //������ ��������
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
                    textBoxRpdGenTargetDir.Text = Path.Combine(Environment.CurrentDirectory, $"���_{DateTime.Now:yyyy-MM-dd}");
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
            }
        }

        private void fastObjectListViewDisciplineListForGeneration_CellToolTipShowing(object sender, ToolTipShowingEventArgs e) {
            if (e.Model is CurriculumDiscipline discipline) {
                if (e.Column.AspectName.Equals("Errors")) {
                    e.Text = string.Join("\r\n", discipline.Errors);
                    e.StandardIcon = ToolTipControl.StandardIcons.Error;
                    e.IsBalloon = true;
                }
            }
        }

        private void fastObjectListViewDisciplineListForGeneration_FormatRow(object sender, FormatRowEventArgs e) {
            if (e.Model is CurriculumDiscipline discipline) {
                if (discipline.Errors?.Any() ?? false) {
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

                if (MessageBox.Show($"�� �������, ��� ������ ������������� ��� ��� ��������� ���������:\r\n  * {disciplineList}\r\n" +
                                    $"� ������� ����������:\r\n{textBoxRpdGenTargetDir.Text}\r\n?\r\n" +
                                    $"����� ��������� ����� � ������� ���������� ����� ������������.\r\n" +
                                    $"����������?",
                                    "��������� ���", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                    var curriculumGroup = comboBoxRpdGenCurriculumGroups.SelectedItem as CurriculumGroup;
                    curriculumGroup.DirectionCode = textBoxRpdGenDirectionCode.Text;
                    curriculumGroup.DirectionName = textBoxRpdGenDirectionName.Text;
                    curriculumGroup.Department = textBoxRpdGenDepartment.Text;
                    curriculumGroup.Profile = textBoxRpdGenProfile.Text;
                    curriculumGroup.FormsOfStudyList = textBoxRpdGenFormsOfStudy.Text;

                    curriculumGroup.CheckedDisciplines = disciplines;
                    var rpdTemplate = Path.Combine(Environment.CurrentDirectory, DIR_TEMPLATES, comboBoxRpdGenTemplates.SelectedItem.ToString());

                    labelRpdGenStatus.Visible = true;
                    Application.UseWaitCursor = true;
                    labelRpdGenStatus.Text = "";

                    var newRpdFiles = App.GenerateRpdFiles(curriculumGroup, rpdTemplate, textBoxRpdGenTargetDir.Text,
                                                           textBoxRpdGenFileNameTemplate.Text,
                                                           (int idx, CurriculumDiscipline discipline) => {
                                                               this.Invoke(new MethodInvoker(() => {
                                                                   labelRpdGenStatus.Text = $"��������� ��� ��� ����������\r\n[{discipline.Name}]\r\n" +
                                                                                            $"({idx} �� {curriculumGroup.CheckedDisciplines.Count})... ";
                                                                   Application.DoEvents();
                                                               }));
                                                           },
                                                           out var errors);

                    Application.UseWaitCursor = false;
                    labelRpdGenStatus.Text += " ���������.";
                    var msg = "";
                    if (newRpdFiles.Count > 0) {
                        msg = $"��������� ��� ���������.\r\n������� ��������� ��� ({newRpdFiles.Count} ��.):\r\n{string.Join("\r\n", newRpdFiles)}";
                    }
                    else {
                        msg = $"��������� ��� ���������.\r\n�������� ������ �� �������.";
                    }
                    if (errors.Any()) {
                        msg += $"\r\n\r\n���������� ������ ({errors.Count} ��.):\r\n{string.Join("\r\n", errors)}";
                    }
                    MessageBox.Show(msg, "��������� ���", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else {
                MessageBox.Show($"��� ��������� ��� ���������� �������������� ������� ������ ������� ������ � �������� ������ ���������� � ������.", "��������� ���",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void label9_Click(object sender, EventArgs e) {

        }

        private void linkLabelRpdGenCompetenceMatrix_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            tabControl1.SelectTab(tabPageCompetenceMatrix);
        }
    }
}
