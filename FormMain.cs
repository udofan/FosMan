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
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace FosMan {
    public partial class FormMain : Form {
        const string DIR_TEMPLATES = "Templates";
        const string DIR_REPORTS = "Reports";
        const string DIR_LOGS = "Logs";

        //string m_matrixFileName = null;
        //string m_rpdHtmlReport = null;

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
                buttonLoadCompetenceMatrix.PerformClick();
                //m_matrixFileName = openFileDialog1.FileName;
            }
        }

        private void buttonLoadCompetenceMatrix_Click(object sender, EventArgs e) {
            LoadCompetenceMatrix(textBoxMatrixFileName.Text, false);
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
                        if ((discipline.Errors?.Any() ?? false) || (discipline.ExtraErrors?.Any() ?? false)) {
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
            var olvColumnDepartmentName = new OLVColumn("�������", "DepartmentName") {
                Width = 100,
                IsEditable = false
            };
            list.Columns.Add(olvColumnDepartmentName);
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
            //�����������
            var olvColumnCompiler = new OLVColumn("�����������", "Compiler") {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnCompiler);
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
            //�������������� ����������
            var olvColumnPrevDisciplines = new OLVColumn("�������������� ����������", "PrevDisciplines") {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnPrevDisciplines);
            //����������� ����������
            var olvColumnNextDisciplines = new OLVColumn("����������� ����������", "NextDisciplines") {
                Width = 200,
                IsEditable = false
            };
            list.Columns.Add(olvColumnNextDisciplines);
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

        void TuneRpdFindAndReplaceList() {
            var olvColFind = new OLVColumn("�����", "FindPattern") {
                Width = 200,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            fastObjectListViewRpdFixFindAndReplaceItems.Columns.Add(olvColFind);
            var olvColReplace = new OLVColumn("�������� ��", "ReplacePattern") {
                Width = 200,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            fastObjectListViewRpdFixFindAndReplaceItems.Columns.Add(olvColReplace);
        }

        void TuneDocPropertiesList() {
            var olvColName = new OLVColumn("��������", "Name") {
                Width = 150,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            fastObjectListViewDocProperties.Columns.Add(olvColName);
            var olvColValue = new OLVColumn("��������", "Value") {
                Width = 150,
                CellEditUseWholeCell = true,
                IsEditable = true
            };
            fastObjectListViewDocProperties.Columns.Add(olvColValue);
        }

        private void FormMain_Load(object sender, EventArgs e) {
            Xceed.Words.NET.Licenser.LicenseKey = "WDN30-W7F00-6RL0S-ERHA";

            App.LoadConfig();

            YaGpt.Init();

            TuneCurriculumList();
            TuneRpdList(fastObjectListViewRpdList);
            TuneDisciplineList(fastObjectListViewDisciplines, true);
            TuneDisciplineList(fastObjectListViewDisciplineListForGeneration, false);
            TuneRpdFindAndReplaceList();
            TuneDocPropertiesList();

            //������ ��������
            var templateDir = Path.Combine(Environment.CurrentDirectory, DIR_TEMPLATES);
            if (Directory.Exists(templateDir)) {
                var files = Directory.GetFiles(templateDir, "*.docx");
                var newTemplates = files.Select(f => Path.GetFileName(f)).Except(comboBoxRpdGenTemplates.Items.Cast<string>());
                if (newTemplates.Any()) {
                    comboBoxRpdGenTemplates.Items.AddRange(newTemplates.ToArray());
                }
            }

            //����������� �������� �� �������
            fastObjectListViewRpdFixFindAndReplaceItems.AddObjects(App.Config.RpdFindAndReplaceItems);
            fastObjectListViewDocProperties.AddObjects(App.Config.RpdFixDocPropertyList);
            textBoxMatrixFileName.Text = App.Config.CompetenceMatrixFileName ?? "";
            ShowHideRpdFixMode(false);
            checkBoxStoreCurriculumList.Checked = App.Config.StoreCurriculumList;
            checkBoxCompetenceMatrixAutoload.Checked = App.Config.CompetenceMatrixAutoload;
            checkBoxStoreRpdList.Checked = App.Config.StoreRpdList;
            checkBoxRpdFixTableOfCompetences.Checked = App.Config.RpdFixTableOfCompetences;
            checkBoxRpdFixSummaryTableOfEduWorks.Checked = App.Config.RpdFixTableOfEduWorks;
            textBoxRpdFixTargetDir.Text = App.Config.RpdFixTargetDir;
            checkBoxRpdFixFillEduWorkTables.Checked = App.Config.RpdFixFillEduWorkTables;
            textBoxRpdGenFileNameTemplate.Text = App.Config.RpdGenFileNameTemplate;
            textBoxRpdGenTargetDir.Text = App.Config.RpdGenTargetDir;
            textBoxRpdFixFileTemplate.Text = App.Config.RpdFixTemplateFileName;
            checkBoxRpdFixByTemplate.Checked = App.Config.RpdFixByTemplate;
            checkBoxRpdFixSetPrevAndNextDisciplines.Checked = App.Config.RpdFixSetPrevAndNextDisciplines;
            checkBoxRpdFixRemoveColorSelection.Checked = App.Config.RpdFixRemoveColorSelections;

            if (!string.IsNullOrEmpty(App.Config.RpdGenTemplate)) {
                if (comboBoxRpdGenTemplates.Items.Contains(App.Config.RpdGenTemplate)) {
                    comboBoxRpdGenTemplates.SelectedItem = App.Config.RpdGenTemplate;
                }
            }
            if (comboBoxRpdGenTemplates.SelectedIndex < 0 && comboBoxRpdGenTemplates.Items.Count > 0) {
                comboBoxRpdGenTemplates.SelectedIndex = 0;
            }

            //�������
            //textBoxMatrixFileName.Text = @"c:\FosMan\�������_�����������\test5.docx";
            //buttonLoadCompetenceMatrix.PerformClick();
            //var files = new string[] {
            //    @"c:\FosMan\�������_�����\������- ����������� 38.03.05 ����� 23.plx.xlsx",
            //    @"c:\FosMan\�������_�����\������-����������� 38.03.05 �� 23.plx.xlsx",
            //    @"c:\FosMan\�������_�����\������-����������� ����� 38.03.05.plx.xlsx"
            //};
            Task.Run(() => {
                this.Invoke(new MethodInvoker(() => {
                    tabControl1.SelectTab(tabPageCompetenceMatrix);
                    if (App.Config.CompetenceMatrixAutoload) {
                        textBoxMatrixFileName.Text = App.Config.CompetenceMatrixFileName;
                        LoadCompetenceMatrix(textBoxMatrixFileName.Text, true);
                    }

                    if (App.Config.CurriculumList?.Any() ?? false) {
                        tabControl1.SelectTab(tabPage�urriculum);
                        LoadCurriculumFiles(App.Config.CurriculumList.ToArray());
                    }
                    if (App.Config.RpdList?.Any() ?? false) {
                        tabControl1.SelectTab(tabPageRpd);
                        LoadRpdFilesAsync(App.Config.RpdList.ToArray());
                    }
                }));
            });
        }

        void StatusMessage(string msg, bool isError = false) {
            toolStripStatusLabel1.Text = msg;
            toolStripStatusLabel1.ForeColor = isError ? Color.Red : SystemColors.ControlText;
        }

        /// <summary>
        /// �������� ������� �����������
        /// </summary>
        /// <param name="fileName"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void LoadCompetenceMatrix(string fileName, bool silent) {
            var loadMatrix = true;

            if (string.IsNullOrEmpty(fileName) || !File.Exists(fileName)) {
                if (!silent) {
                    MessageBox.Show("���� ������� ����������� �� �����.", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                loadMatrix = false;
            }

            if (loadMatrix && App.CompetenceMatrix != null && App.CompetenceMatrix.IsLoaded) {
                if (!silent) {
                    if (MessageBox.Show("������� ��� ���������.\r\n��������?", "��������", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation) != DialogResult.Yes) {
                        loadMatrix = false;
                    }
                }
            }

            if (loadMatrix) {
                var matrix = CompetenceMatrix.LoadFromFile(fileName);
                var errLog = new ConcurrentDictionary<string, List<string>>();

                if (matrix.IsLoaded) {
                    App.SetCompetenceMatrix(matrix);
                    if (matrix.Errors.Any()) {
                        errLog.TryAdd(fileName, matrix.Errors);
                    }
                    //m_matrixFileName = fileName;

                    ShowMatrix(matrix);
                    //MessageBox.Show("������� ����������� ������� ���������.", "��������", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (errLog.Any()) {
                    WriteErrorLog(errLog, "�������_�����������");
                    //var logDir = Path.Combine(Environment.CurrentDirectory, DIR_LOGS);
                    //if (!Directory.Exists(logDir)) {
                    //    Directory.CreateDirectory(logDir);
                    //}
                    //var dt = DateTime.Now;
                    //var reportFile = Path.Combine(logDir, $"�����������.{dt:yyyy-MM-dd_HH-mm-ss}.log");
                    //var lines = new List<string>() { fileName };
                    //lines.AddRange(matrix.Errors);
                    //File.WriteAllLines(reportFile, lines, Encoding.UTF8);

                    //if (!silent) {
                    //    StatusMessage($"�� ����� �������� ���������� ������ ({matrix.Errors}). ��� ��������� � ������� {reportFile}", true);
                    //    //var errorList = string.Join("\r\n", matrix.Errors);
                    //    //MessageBox.Show($"�� ����� �������� ���������� ������:\r\n{errorList}\r\n\r\n��� ��������� � ������� {reportFile}.", "�������� �������",
                    //      //              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    //}
                }
                else {
                    StatusMessage("������� ����������� ������� ���������");
                }
            }
        }

        /// <summary>
        /// �������� �� �� ������ ���� ������
        /// </summary>
        /// <param name="files"></param>
        void LoadCurriculumFiles(string[] files) {
            labelExcelFileLoading.Visible = true;
            Application.UseWaitCursor = true;
            labelExcelFileLoading.Text = "";

            //var report = new List<string>();
            var errLog = new ConcurrentDictionary<string, List<string>>();
            var idx = 1;
            foreach (var file in files) {
                labelExcelFileLoading.Text = $"�������� ������ ({idx} �� {files.Length})...";
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
                labelExcelFileLoading.Text += " ���������.";
            }
            if (errLog.Any()) {
                var logFile = WriteErrorLog(errLog, "��");

                //MessageBox.Show($"�� ����� �������� ���������� ������.\r\n��� ��������� � ������� {reportFile}.", "��������",
                //              MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else {
                StatusMessage($"������� ��������� {files.Length} ����(�,��)");
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

                LoadRpdFilesAsync(files);
            }
        }

        /// <summary>
        /// ����� ������� ������
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
                StatusMessage($"�� ����� �������� ���������� ������ ({errCount} ��.). ��. ������ [{logFile}]", true);
            }
            catch (Exception ex) {
            }

            return logFile;
        }

        /// <summary>
        /// �������� ���-������
        /// </summary>
        /// <param name="files"></param>
        private async Task LoadRpdFilesAsync(string[] files) {
            labelLoadRpd.Visible = true;
            Application.UseWaitCursor = true;
            labelLoadRpd.Text = "";

            var errLog = new ConcurrentDictionary<string, List<string>>();
            var idx = 0;

            StatusMessage("��������� �����...");

            //RpdAddEventHandler OnRpdAdd = (rpd) => {
            //    this.BeginInvoke(new MethodInvoker(() => {
            //        fastObjectListViewRpdList.AddObject(rpd);
            //        labelLoadRpd.Text = $"�������� ������ ({idx} �� {files.Length})...";
            //    }));
            //};
            //App.RpdAdd += OnRpdAdd;

            var tasks = new List<Task>();

            foreach (var file in files) {
                tasks.Add(Task.Run(() => {
                    Interlocked.Increment(ref idx);

                    if (File.Exists(file)) {
                        var rpd = Rpd.LoadFromFile(file);
                        App.AddRpd(rpd);

                        this.Invoke(new MethodInvoker(() => {
                            fastObjectListViewRpdList.BeginUpdate();
                            var selectedObjects = fastObjectListViewRpdList.SelectedObjects;
                            fastObjectListViewRpdList.AddObject(rpd);
                            fastObjectListViewRpdList.EnsureModelVisible(rpd);
                            selectedObjects.Add(rpd);
                            fastObjectListViewRpdList.SelectedObjects = selectedObjects;
                            fastObjectListViewRpdList.EndUpdate();

                            var msg = $"�������� ������ ({idx} �� {files.Length})...";
                            labelLoadRpd.Text = msg;
                            StatusMessage(msg);
                            //if (idx == files.Length) {
                            //    labelLoadRpd.Text += " ���������.";
                            //}
                            Application.DoEvents();
                        }));

                        if (rpd.Errors.Any()) {
                            errLog.TryAdd(file, rpd.Errors);
                        }
                    }
                }));
            }

            await Task.WhenAll(tasks.ToArray());

            labelLoadRpd.Text += " ���������.";

            if (errLog.Any()) {
                var logFile = WriteErrorLog(errLog, "���");

                //var errCount = errLog.Sum(r => r.Value.Count);
                //StatusMessage($"�� ����� �������� ���������� ������ ({errCount}). ��� ��������� � ������� {logFile}", true);
                //MessageBox.Show($"�� ����� �������� ���������� ������.\r\n��� ��������� � ������� {reportFile}.", "��������",
                //              MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else {
                StatusMessage($"������� ��������� {files.Length} ����(�,��)");
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
            var rpdList = fastObjectListViewRpdList.SelectedObjects?.Cast<Rpd>().ToList();
            if (rpdList.Any()) {
                var targetDir = textBoxRpdGenTargetDir.Text;
                if (string.IsNullOrEmpty(targetDir)) {
                    targetDir = Path.Combine(Environment.CurrentDirectory, $"������������_���_{DateTime.Now:yyyy-MM-dd}");
                }
                App.CheckRdp(rpdList, out var report);

                AddReport("�������� ���", report);
            }
            else {
                MessageBox.Show("���������� �������� �����, ������� ��������� ���������.", "����������� ���", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// �������� ������� � �������
        /// </summary>
        /// <param name="name"></param>
        /// <param name="html"></param>
        async void AddReport(string name, string html) {
            if (!string.IsNullOrEmpty(html)) {
                var tabCount = tabControlReports.TabPages.Cast<TabPage>().Select(t => t.Name.StartsWith(name, StringComparison.CurrentCultureIgnoreCase)).Count();
                name += $" ({tabCount + 1})  [x]";

                var tabPageNewReport = new TabPage(name);

                var webView = new WebView2();
                await webView.EnsureCoreWebView2Async();
                webView.NavigateToString(html);
                tabPageNewReport.Controls.Add(webView);
                webView.Dock = DockStyle.Fill;

                tabPageNewReport.Tag = html;    //���������� �� ������ ����������
                tabControlReports.TabPages.Add(tabPageNewReport);
                tabControlReports.SelectTab(tabPageNewReport);

                tabControl1.SelectTab(tabPageReports);
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
                var reportFile = Path.Combine(repDir, $"�����.{repName}.{dt:yyyy-MM-dd_HH-mm-ss}.html");
                File.WriteAllText(reportFile, htmlReport, Encoding.UTF8);
                MessageBox.Show($"����� �������� � ����\r\n{reportFile}", "���������� ������",
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

                if (string.IsNullOrEmpty(textBoxRpdGenTargetDir.Text)) {
                    textBoxRpdGenTargetDir.Text = Path.Combine(Environment.CurrentDirectory, $"���_{DateTime.Now:yyyy-MM-dd}");
                    textBoxRpdGenTargetDir.SelectionStart = textBoxRpdGenTargetDir.Text.Length - 1;
                }
            }
            if (tabControl1.SelectedTab == tabPageRpd) {
                if (string.IsNullOrEmpty(textBoxRpdFixTargetDir.Text)) {
                    textBoxRpdFixTargetDir.Text = Path.Combine(Environment.CurrentDirectory, $"������������_���_{DateTime.Now:yyyy-MM-dd}");
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
                    curriculumGroup.FSES = textBoxRpdGenFSES.Text;

                    curriculumGroup.CheckedDisciplines = disciplines;
                    var rpdTemplate = Path.Combine(Environment.CurrentDirectory, DIR_TEMPLATES, comboBoxRpdGenTemplates.SelectedItem.ToString());

                    labelRpdGenStatus.Visible = true;
                    Application.UseWaitCursor = true;
                    labelRpdGenStatus.Text = "";
                    var targetDir = textBoxRpdGenTargetDir.Text;
                    if (string.IsNullOrEmpty(targetDir)) {
                        targetDir = Path.Combine(Environment.CurrentDirectory, $"���_{DateTime.Now:yyyy-MM-dd}");
                    }
                    var newRpdFiles = App.GenerateRpdFiles(curriculumGroup, rpdTemplate, targetDir,
                                                           textBoxRpdGenFileNameTemplate.Text,
                                                           (int idx, CurriculumDiscipline discipline) => {
                                                               this.Invoke(new MethodInvoker(() => {
                                                                   labelRpdGenStatus.Text = $"��������� ��� ��� ����������\r\n[{discipline.Name}]\r\n" +
                                                                                            $"({idx} �� {curriculumGroup.CheckedDisciplines.Count})... ";
                                                                   Application.DoEvents();
                                                               }));
                                                           },
                                                           checkBoxApplyLoadedRpd.Checked,
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

        private void linkLabelRpdPage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            tabControl1.SelectTab(tabPageRpd);
        }

        void ShowHideRpdFixMode(bool show) {
            var maxDist = 280;
            var delta = 10;

            fastObjectListViewRpdList.BeginUpdate();

            if (show) {
                while (splitContainer1.SplitterDistance < maxDist) {
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
            fastObjectListViewRpdList.EndUpdate();
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
                var ret = MessageBox.Show($"�� �������, ��� ������ ������� ���������� �������� ({fastObjectListViewRpdFixFindAndReplaceItems.SelectedObjects.Count} ��.)?",
                    "��������", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
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
            if (MessageBox.Show("�� �������, ��� ������ �������� ������ ����������� ������� ������?",
                "��������", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
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

        private void tabControlReports_DrawItem(object sender, DrawItemEventArgs e) {
            //e.Graphics.DrawString("x", e.Font, Brushes.Black, e.Bounds.Right - 15, e.Bounds.Top + 4);
            //e.Graphics.DrawString(this.tabControl1.TabPages[e.Index].Text, e.Font, Brushes.Black, e.Bounds.Left + 12, e.Bounds.Top + 4);
            //e.DrawFocusRectangle();
        }

        private void tabControlReports_MouseDown(object sender, MouseEventArgs e) {
            Rectangle r = tabControlReports.GetTabRect(tabControlReports.SelectedIndex);
            Rectangle closeButton = new Rectangle(r.Right - 22, r.Top + 3, 15, 9);
            if (closeButton.Contains(e.Location)) {
                if (MessageBox.Show("������� �������?", "��������", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                    tabControlReports.TabPages.Remove(tabControlReports.SelectedTab);
                }
            }
        }

        private void tabControlReports_MouseMove(object sender, MouseEventArgs e) {
            //NativeMethods.TCHITTESTINFO HTI = new NativeMethods.TCHITTESTINFO(tabControl.PointToClient(Cursor.Position));
            //int tabID = NativeMethods.SendMessage(tabControl.Handle, NativeMethods.TCM_HITTEST, IntPtr.Zero, ref HTI);
            //return tabID == -1 ? null : tabControl.TabPages[tabID];

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
                    targetDir = Path.Combine(Environment.CurrentDirectory, $"������������_���_{DateTime.Now:yyyy-MM-dd}");
                }
                App.FixRpdFiles(rpdList, targetDir, out var htmlReport);

                AddReport("����������� ���", htmlReport);
            }
            else {
                MessageBox.Show("���������� �������� �����, ������� ��������� ���������.", "����������� ���", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
            var initDir = Directory.Exists(buttonRpdFixSelectTargetDir.Text) ? buttonRpdFixSelectTargetDir.Text : Environment.CurrentDirectory;
            folderBrowserDialogSelectDir.InitialDirectory = initDir;

            if (folderBrowserDialogSelectDir.ShowDialog() == DialogResult.OK) {
                buttonRpdFixSelectTargetDir.Text = folderBrowserDialogSelectDir.SelectedPath;
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
            if (MessageBox.Show("�� �������, ��� ������ �������� ������ ����������� ���?",
                                "��������", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
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
                MessageBox.Show($"��� ��������� ��� �� ������� ���������� ���������:\n\n" +
                                $"{string.Join("\n", rpdList.Select(r => r.DisciplineName))}\n\n" +
                                $"��� ���� �������� � ������.", "��������!",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void fastObjectListViewDisciplineListForGeneration_ItemChecked(object sender, ItemCheckedEventArgs e) {
            groupBoxRpdGenDisciplineList.Text = $"3. ����� ��������� [{fastObjectListViewDisciplineListForGeneration.CheckedObjects?.Count ?? 0}]";
        }

        private void fastObjectListViewRpdList_ItemActivate(object sender, EventArgs e) {
            if (fastObjectListViewRpdList.FocusedObject is Rpd rpd) {
                var p = new Process();
                p.StartInfo = new ProcessStartInfo(rpd.SourceFileName) {
                    UseShellExecute = true
                };
                p.Start();
            }
        }

        private async void buttonYaGptSendQuestion_Click(object sender, EventArgs e) {
            var temp = double.Parse(textBoxYaGptTemp.Text);
            var ret = await YaGpt.TextGenerationAsync(textBoxYaGptSystemText.Text, textBoxYaGptUserText.Text, temp);

            textBoxYaGptAnswer.Text = ret.text ?? ret.error?.message;
        }

        private void button1_Click_1(object sender, EventArgs e) {
            textBoxYaGptSystemText.Text = "�� - ��������� ������������";
            textBoxYaGptUserText.Text = "����� ���� � ������ � ���������� \"����������� �������� �����\"? �� ��������� �������� Markdown!";
        }

        private void checkBoxRpdFixFillEduWorkTables_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixFillEduWorkTables = checkBoxRpdFixFillEduWorkTables.Checked;
            App.SaveConfig();
        }

        private void checkBoxSetPrevAndNextDisciplines_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixSetPrevAndNextDisciplines = checkBoxRpdFixSetPrevAndNextDisciplines.Checked;
            App.SaveConfig();
        }

        private void buttonRpdSaveToDb_Click(object sender, EventArgs e) {
            var sw = Stopwatch.StartNew();

            App.AddLoadedRpdToStore();

            StatusMessage($"� ���� ��������� ��� ({App.RpdList.Count} ��.) ({sw.Elapsed}). ����� � ����� ���: {App.Store.RpdDic.Count} ��.");
        }

        private void buttonRpdLoadFromDb_Click(object sender, EventArgs e) {
            if (MessageBox.Show($"�� �������, ��� ������ ��������� � ������ ��� �� ����� ({App.Store.RpdDic.Count} ��.)?", "�������� ���",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) == DialogResult.Yes) {
                if (App.Store?.RpdDic?.Any() ?? false) {
                    fastObjectListViewRpdList.BeginUpdate();
                    var selectedObjects = fastObjectListViewRpdList.SelectedObjects;

                    foreach (var rpd in App.Store.RpdDic.Values) {
                        App.AddRpd(rpd);

                        fastObjectListViewRpdList.AddObject(rpd);
                        fastObjectListViewRpdList.EnsureModelVisible(rpd);
                        selectedObjects.Add(rpd);
                    }
                    fastObjectListViewRpdList.SelectedObjects = selectedObjects;
                    fastObjectListViewRpdList.EndUpdate();
                }
            }
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

                    //var msg = $"�������� ������ ({idx} �� {files.Length})...";
                    //labelLoadRpd.Text = msg;
                    //StatusMessage(msg);
                    //if (idx == files.Length) {
                    //    labelLoadRpd.Text += " ���������.";
                    //}
                    //Debug.WriteLine($"{idx}: inside form.Invoke");
                    StatusMessage($"���������� ��� ({idx} �� {rpdList.Count}) (��������� ������: {sw.Elapsed}, �����: {swMain.Elapsed})...");
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
                StatusMessage("���������� �������� �� ��������� ������� ���������...");
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
                    //StatusMessage($"���������� �������� �� ��������� ������� ��������� (��������� {i} �� {rpdList.Count}) [{sw.Elapsed}]...");
                }
                */
                await GenerateRelatedDisciplines(rpdList, true);
                StatusMessage($"������� �� ��������� ������� ��������� ��������� ({sw.Elapsed}).");
            }
            else {
                MessageBox.Show("���������� �������� �����, � ������� ��������� ��������� �������.", "��������� ������ ���������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
                StatusMessage("���������� �������� �� ��������� ������� ���������...");
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
                //    StatusMessage($"���������� �������� �� ��������� ������� ��������� (��������� {i} �� {rpdList.Count}) [{sw.Elapsed}]...");
                //}
                await GenerateRelatedDisciplines(rpdList, false);
                StatusMessage($"������� �� ��������� ������� ��������� ��������� ({sw.Elapsed}).");
            }
            else {
                MessageBox.Show("���������� �������� �����, � ������� ��������� ��������� �������.", "��������� ������ ���������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void checkBoxRpdFixResetColorSelection_CheckedChanged(object sender, EventArgs e) {
            App.Config.RpdFixRemoveColorSelections = checkBoxRpdFixRemoveColorSelection.Checked;
            App.SaveConfig();
        }
    }
}
