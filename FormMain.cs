using BrightIdeasSoftware;

namespace FosMan {
    public partial class FormMain : Form {
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

            if (loadMatrix && CompetenceMatrix.IsLoaded) {
                if (MessageBox.Show("Матрица уже загружена.\r\nОбновить?", "Внимание", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation) != DialogResult.Yes) {
                    loadMatrix = false;
                }
            }

            if (loadMatrix) {
                CompetenceMatrix.LoadFromFile(textBoxMatrixFileName.Text, out var errors);

                if (CompetenceMatrix.IsLoaded) {
                    m_matrixFileName = textBoxMatrixFileName.Text;

                    ShowMatrix(errors);
                    MessageBox.Show("Матрица компетенций успешно загружена.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else {
                    var errorList = string.Join("\r\n", errors);
                    MessageBox.Show($"Обнаружены ошибки:\r\n{errorList}", "Загрузка матрицы", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
        }

        /// <summary>
        /// Покажем матрицу компетенций в WebView2
        /// </summary>
        /// <param name="errors"></param>
        async void ShowMatrix(List<string> errors) {
            var html = "<html><body>";
            var tdStyle = " style='border: 1px solid; vertical-align: top'";

            //если есть ошибки
            if (errors.Any()) {
                html += "<div style='color: red'><b>ОШИБКИ:</b></div>";
                html += string.Join("", errors.Select(e => $"<div style='color: red'>{e}</div>"));
            }

            //формирование матрицы
            html += "<table style='border: 1px solid'>";
            html += $"<tr><th {tdStyle}><b>Код и наименование компетенций</b></th>" +
                        $"<th {tdStyle}><b>Коды и индикаторы достижения компетенций\r\n</b></tdh>" +
                        $"<th {tdStyle}><b>Коды и результаты обучения</b></td></th></tr>";
            foreach (var item in CompetenceMatrix.Items) {
                html += "<tr>";
                html += $"<td {tdStyle} rowspan='{item.Achievements.Count}'><b>{item.Code}</b>. {item.Title}</td>";
                var firstAchi = true;
                foreach (var achi in item.Achievements) {
                    if (!firstAchi) {
                        html += "<tr>";
                    }
                    html += $"<td {tdStyle}>{achi.Code}. {achi.Indicator}</td>";
                    html += $"<td {tdStyle}>{achi.ResultCode}:<br />{achi.ResultDescription}</td>";
                    if (!firstAchi) {
                        html += "</tr>";
                    }
                    firstAchi = false;
                }
                html += "</tr>";
            }
            html += "</table>";

            html += "</body></html>";

            await webView21.EnsureCoreWebView2Async();
            webView21.NavigateToString(html);
        }

        void TuneCurriculumList() {
            var olvColumnCurriculumIndex = new OLVColumn("Код", "ProgramCode") {
                Width = 80,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumIndex);
            var olvColumnCurriculumName = new OLVColumn("Программа", "ProgramName") {
                Width = 200,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumName);
            var olvColumnCurriculumProfile = new OLVColumn("Программа", "Profile") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumProfile);
            var olvColumnCurriculumDepartment = new OLVColumn("Программа", "Department") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumDepartment);
            var olvColumnCurriculumFaculty = new OLVColumn("Программа", "Faculty") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumFaculty);
            var olvColumnCurriculumAcademicYear = new OLVColumn("Программа", "AcademicYear") {
                Width = 100,
                IsEditable = false
            };
            fastObjectListViewCurricula.Columns.Add(olvColumnCurriculumAcademicYear);
            var olvColumnCurriculumFormOfStudy = new OLVColumn("Программа", "FormOfStudy") {
                Width = 100,
                IsEditable = false,
                AspectGetter = (x) => {
                    var value = x?.ToString();
                    var curriculum = x as Curriculum;
                    if (curriculum != null) {
                        if (curriculum.FormOfStudy == EFormOfStudy.FullTime) {
                            value = "очная";
                        }
                        else if (curriculum.FormOfStudy == EFormOfStudy.PartTime) {
                            value = "заочная";
                        }
                        else if (curriculum.FormOfStudy == EFormOfStudy.FullTimeAndPartTime) {
                            value = "очно-заочная";
                        }
                        else {
                            value = "НЕИЗВЕСТНО";
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
        }

        private void FormMain_Load(object sender, EventArgs e) {
            TuneCurriculumList();
            TuneDisciplineList();
        }

        private void buttonSelectExcelFiles_Click(object sender, EventArgs e) {
            if (openFileDialog2.ShowDialog(this) == DialogResult.OK) {
                var files = openFileDialog2.FileNames.Where(x => !Curricula.HasFile(x)).ToArray();

                labelExcelFileLoading.Visible = true;
                Application.UseWaitCursor = true;
                labelExcelFileLoading.Text = "";

                var idx = 1;
                foreach (var file in files) {
                    labelExcelFileLoading.Text = $"Загрузка файлов ({idx} из {files.Length})...";
                    var curriculum = Curriculum.LoadFromFile(file);
                    Curricula.AddCurriculum(curriculum);

                    fastObjectListViewCurricula.AddObject(curriculum);
                    idx++;
                    Application.DoEvents();
                }
                if (idx > 1) {
                    labelExcelFileLoading.Text += " завершено.";
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
    }
}
