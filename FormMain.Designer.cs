namespace FosMan {
    partial class FormMain {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            buttonSelectFosDir = new Button();
            textBoxFosDir = new TextBox();
            fastObjectListViewFosList = new BrightIdeasSoftware.FastObjectListView();
            tabControl1 = new TabControl();
            tabPageCompetenceMatrix = new TabPage();
            webView21 = new Microsoft.Web.WebView2.WinForms.WebView2();
            buttonLoadCompetenceMatrix = new Button();
            buttonSelectMatrixFile = new Button();
            label2 = new Label();
            textBoxMatrixFileName = new TextBox();
            tabPageСurriculum = new TabPage();
            groupBoxDisciplines = new GroupBox();
            fastObjectListViewDisciplines = new BrightIdeasSoftware.FastObjectListView();
            labelExcelFileLoading = new Label();
            fastObjectListViewCurricula = new BrightIdeasSoftware.FastObjectListView();
            buttonSelectExcelFiles = new Button();
            tabPage1 = new TabPage();
            buttonRpdCheck = new Button();
            labelLoadRpd = new Label();
            fastObjectListViewRpdList = new BrightIdeasSoftware.FastObjectListView();
            buttonSelectRpdFiles = new Button();
            tabPage2 = new TabPage();
            label1 = new Label();
            tabPageRpdCheck = new TabPage();
            buttonSaveRpdReport = new Button();
            webView2RpdReport = new Microsoft.Web.WebView2.WinForms.WebView2();
            tabPageRpdGeneration = new TabPage();
            labelRpdGenStatus = new Label();
            label10 = new Label();
            textBoxRpdGenDepartment = new TextBox();
            label9 = new Label();
            textBoxRpdGenProfile = new TextBox();
            textBoxRpdGenDirectionName = new TextBox();
            label8 = new Label();
            textBoxRpdGenDirectionCode = new TextBox();
            label7 = new Label();
            groupBox2 = new GroupBox();
            comboBoxRpdGenTemplates = new ComboBox();
            buttonSelectRpdTargetDir = new Button();
            label4 = new Label();
            textBoxRpdGenTargetDir = new TextBox();
            label6 = new Label();
            textBoxRpdGenFileNameTemplate = new TextBox();
            label5 = new Label();
            buttonGenerate = new Button();
            groupBox1 = new GroupBox();
            fastObjectListViewDisciplineListForGeneration = new BrightIdeasSoftware.FastObjectListView();
            label3 = new Label();
            comboBoxRpdGenCurriculumGroups = new ComboBox();
            openFileDialog1 = new OpenFileDialog();
            openFileDialog2 = new OpenFileDialog();
            openFileDialog3 = new OpenFileDialog();
            openFileDialog4 = new OpenFileDialog();
            openFileDialogRpdTemplate = new OpenFileDialog();
            folderBrowserDialogRpdTargetDir = new FolderBrowserDialog();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewFosList).BeginInit();
            tabControl1.SuspendLayout();
            tabPageCompetenceMatrix.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)webView21).BeginInit();
            tabPageСurriculum.SuspendLayout();
            groupBoxDisciplines.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewDisciplines).BeginInit();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewCurricula).BeginInit();
            tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewRpdList).BeginInit();
            tabPage2.SuspendLayout();
            tabPageRpdCheck.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)webView2RpdReport).BeginInit();
            tabPageRpdGeneration.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewDisciplineListForGeneration).BeginInit();
            SuspendLayout();
            // 
            // buttonSelectFosDir
            // 
            buttonSelectFosDir.Location = new Point(434, 20);
            buttonSelectFosDir.Name = "buttonSelectFosDir";
            buttonSelectFosDir.Size = new Size(75, 23);
            buttonSelectFosDir.TabIndex = 0;
            buttonSelectFosDir.Text = "...";
            buttonSelectFosDir.UseVisualStyleBackColor = true;
            // 
            // textBoxFosDir
            // 
            textBoxFosDir.Location = new Point(123, 21);
            textBoxFosDir.Name = "textBoxFosDir";
            textBoxFosDir.Size = new Size(305, 22);
            textBoxFosDir.TabIndex = 1;
            // 
            // fastObjectListViewFosList
            // 
            fastObjectListViewFosList.GridLines = true;
            fastObjectListViewFosList.Location = new Point(21, 60);
            fastObjectListViewFosList.Name = "fastObjectListViewFosList";
            fastObjectListViewFosList.ShowGroups = false;
            fastObjectListViewFosList.Size = new Size(496, 404);
            fastObjectListViewFosList.TabIndex = 0;
            fastObjectListViewFosList.View = View.Details;
            fastObjectListViewFosList.VirtualMode = true;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPageCompetenceMatrix);
            tabControl1.Controls.Add(tabPageСurriculum);
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPageRpdCheck);
            tabControl1.Controls.Add(tabPageRpdGeneration);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1047, 644);
            tabControl1.TabIndex = 1;
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
            // 
            // tabPageCompetenceMatrix
            // 
            tabPageCompetenceMatrix.Controls.Add(webView21);
            tabPageCompetenceMatrix.Controls.Add(buttonLoadCompetenceMatrix);
            tabPageCompetenceMatrix.Controls.Add(buttonSelectMatrixFile);
            tabPageCompetenceMatrix.Controls.Add(label2);
            tabPageCompetenceMatrix.Controls.Add(textBoxMatrixFileName);
            tabPageCompetenceMatrix.Location = new Point(4, 23);
            tabPageCompetenceMatrix.Name = "tabPageCompetenceMatrix";
            tabPageCompetenceMatrix.Padding = new Padding(3);
            tabPageCompetenceMatrix.Size = new Size(1039, 617);
            tabPageCompetenceMatrix.TabIndex = 2;
            tabPageCompetenceMatrix.Text = "Матрица компетенций";
            tabPageCompetenceMatrix.UseVisualStyleBackColor = true;
            // 
            // webView21
            // 
            webView21.AllowExternalDrop = true;
            webView21.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            webView21.CreationProperties = null;
            webView21.DefaultBackgroundColor = Color.White;
            webView21.Location = new Point(8, 59);
            webView21.Name = "webView21";
            webView21.Size = new Size(1023, 550);
            webView21.TabIndex = 4;
            webView21.ZoomFactor = 1D;
            // 
            // buttonLoadCompetenceMatrix
            // 
            buttonLoadCompetenceMatrix.Location = new Point(676, 14);
            buttonLoadCompetenceMatrix.Name = "buttonLoadCompetenceMatrix";
            buttonLoadCompetenceMatrix.Size = new Size(75, 23);
            buttonLoadCompetenceMatrix.TabIndex = 3;
            buttonLoadCompetenceMatrix.Text = "Загрузить";
            buttonLoadCompetenceMatrix.UseVisualStyleBackColor = true;
            buttonLoadCompetenceMatrix.Click += buttonLoadCompetenceMatrix_Click;
            // 
            // buttonSelectMatrixFile
            // 
            buttonSelectMatrixFile.Location = new Point(603, 14);
            buttonSelectMatrixFile.Name = "buttonSelectMatrixFile";
            buttonSelectMatrixFile.Size = new Size(33, 23);
            buttonSelectMatrixFile.TabIndex = 2;
            buttonSelectMatrixFile.Text = "...";
            buttonSelectMatrixFile.UseVisualStyleBackColor = true;
            buttonSelectMatrixFile.Click += button1_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(20, 18);
            label2.Name = "label2";
            label2.Size = new Size(41, 14);
            label2.TabIndex = 1;
            label2.Text = "Файл:";
            // 
            // textBoxMatrixFileName
            // 
            textBoxMatrixFileName.Location = new Point(67, 15);
            textBoxMatrixFileName.Name = "textBoxMatrixFileName";
            textBoxMatrixFileName.Size = new Size(530, 22);
            textBoxMatrixFileName.TabIndex = 0;
            textBoxMatrixFileName.Text = "c:\\FosMan\\Матрицы_компетенций\\test3.docx";
            // 
            // tabPageСurriculum
            // 
            tabPageСurriculum.Controls.Add(groupBoxDisciplines);
            tabPageСurriculum.Controls.Add(labelExcelFileLoading);
            tabPageСurriculum.Controls.Add(fastObjectListViewCurricula);
            tabPageСurriculum.Controls.Add(buttonSelectExcelFiles);
            tabPageСurriculum.Location = new Point(4, 24);
            tabPageСurriculum.Name = "tabPageСurriculum";
            tabPageСurriculum.Padding = new Padding(3);
            tabPageСurriculum.Size = new Size(1039, 616);
            tabPageСurriculum.TabIndex = 3;
            tabPageСurriculum.Text = "Учебные планы";
            tabPageСurriculum.UseVisualStyleBackColor = true;
            // 
            // groupBoxDisciplines
            // 
            groupBoxDisciplines.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBoxDisciplines.Controls.Add(fastObjectListViewDisciplines);
            groupBoxDisciplines.Location = new Point(8, 379);
            groupBoxDisciplines.Name = "groupBoxDisciplines";
            groupBoxDisciplines.Size = new Size(1023, 233);
            groupBoxDisciplines.TabIndex = 3;
            groupBoxDisciplines.TabStop = false;
            groupBoxDisciplines.Text = "Дисциплины программы";
            // 
            // fastObjectListViewDisciplines
            // 
            fastObjectListViewDisciplines.Dock = DockStyle.Fill;
            fastObjectListViewDisciplines.FullRowSelect = true;
            fastObjectListViewDisciplines.GridLines = true;
            fastObjectListViewDisciplines.HeaderWordWrap = true;
            fastObjectListViewDisciplines.Location = new Point(3, 18);
            fastObjectListViewDisciplines.Name = "fastObjectListViewDisciplines";
            fastObjectListViewDisciplines.ShowGroups = false;
            fastObjectListViewDisciplines.Size = new Size(1017, 212);
            fastObjectListViewDisciplines.TabIndex = 4;
            fastObjectListViewDisciplines.UseFilterIndicator = true;
            fastObjectListViewDisciplines.UseFiltering = true;
            fastObjectListViewDisciplines.UseHotItem = true;
            fastObjectListViewDisciplines.UseTranslucentHotItem = true;
            fastObjectListViewDisciplines.UseTranslucentSelection = true;
            fastObjectListViewDisciplines.View = View.Details;
            fastObjectListViewDisciplines.VirtualMode = true;
            fastObjectListViewDisciplines.CellToolTipShowing += fastObjectListViewDisciplines_CellToolTipShowing;
            fastObjectListViewDisciplines.FormatRow += fastObjectListViewDisciplines_FormatRow;
            // 
            // labelExcelFileLoading
            // 
            labelExcelFileLoading.AutoSize = true;
            labelExcelFileLoading.Location = new Point(200, 23);
            labelExcelFileLoading.Name = "labelExcelFileLoading";
            labelExcelFileLoading.Size = new Size(66, 14);
            labelExcelFileLoading.TabIndex = 2;
            labelExcelFileLoading.Text = "Загрузка...";
            labelExcelFileLoading.Visible = false;
            // 
            // fastObjectListViewCurricula
            // 
            fastObjectListViewCurricula.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            fastObjectListViewCurricula.FullRowSelect = true;
            fastObjectListViewCurricula.GridLines = true;
            fastObjectListViewCurricula.Location = new Point(8, 57);
            fastObjectListViewCurricula.Name = "fastObjectListViewCurricula";
            fastObjectListViewCurricula.ShowGroups = false;
            fastObjectListViewCurricula.Size = new Size(1023, 316);
            fastObjectListViewCurricula.TabIndex = 1;
            fastObjectListViewCurricula.UseFilterIndicator = true;
            fastObjectListViewCurricula.UseFiltering = true;
            fastObjectListViewCurricula.UseHotItem = true;
            fastObjectListViewCurricula.UseTranslucentHotItem = true;
            fastObjectListViewCurricula.UseTranslucentSelection = true;
            fastObjectListViewCurricula.View = View.Details;
            fastObjectListViewCurricula.VirtualMode = true;
            fastObjectListViewCurricula.CellToolTipShowing += fastObjectListViewCurricula_CellToolTipShowing;
            fastObjectListViewCurricula.FormatRow += fastObjectListViewCurricula_FormatRow;
            fastObjectListViewCurricula.ItemActivate += fastObjectListViewCurricula_ItemActivate;
            // 
            // buttonSelectExcelFiles
            // 
            buttonSelectExcelFiles.Location = new Point(23, 19);
            buttonSelectExcelFiles.Name = "buttonSelectExcelFiles";
            buttonSelectExcelFiles.Size = new Size(143, 23);
            buttonSelectExcelFiles.TabIndex = 0;
            buttonSelectExcelFiles.Text = "Выбор файлов...";
            buttonSelectExcelFiles.UseVisualStyleBackColor = true;
            buttonSelectExcelFiles.Click += buttonSelectExcelFiles_Click;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(buttonRpdCheck);
            tabPage1.Controls.Add(labelLoadRpd);
            tabPage1.Controls.Add(fastObjectListViewRpdList);
            tabPage1.Controls.Add(buttonSelectRpdFiles);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(1039, 616);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "РПД";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // buttonRpdCheck
            // 
            buttonRpdCheck.Location = new Point(582, 18);
            buttonRpdCheck.Name = "buttonRpdCheck";
            buttonRpdCheck.Size = new Size(100, 23);
            buttonRpdCheck.TabIndex = 6;
            buttonRpdCheck.Text = "Проверить";
            buttonRpdCheck.UseVisualStyleBackColor = true;
            buttonRpdCheck.Click += buttonRpdCheck_Click;
            // 
            // labelLoadRpd
            // 
            labelLoadRpd.AutoSize = true;
            labelLoadRpd.Location = new Point(198, 22);
            labelLoadRpd.Name = "labelLoadRpd";
            labelLoadRpd.Size = new Size(66, 14);
            labelLoadRpd.TabIndex = 5;
            labelLoadRpd.Text = "Загрузка...";
            labelLoadRpd.Visible = false;
            // 
            // fastObjectListViewRpdList
            // 
            fastObjectListViewRpdList.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            fastObjectListViewRpdList.FullRowSelect = true;
            fastObjectListViewRpdList.GridLines = true;
            fastObjectListViewRpdList.Location = new Point(8, 57);
            fastObjectListViewRpdList.Name = "fastObjectListViewRpdList";
            fastObjectListViewRpdList.ShowGroups = false;
            fastObjectListViewRpdList.Size = new Size(1023, 554);
            fastObjectListViewRpdList.TabIndex = 4;
            fastObjectListViewRpdList.UseFilterIndicator = true;
            fastObjectListViewRpdList.UseFiltering = true;
            fastObjectListViewRpdList.UseHotItem = true;
            fastObjectListViewRpdList.UseTranslucentHotItem = true;
            fastObjectListViewRpdList.UseTranslucentSelection = true;
            fastObjectListViewRpdList.View = View.Details;
            fastObjectListViewRpdList.VirtualMode = true;
            fastObjectListViewRpdList.CellToolTipShowing += fastObjectListViewRpdList_CellToolTipShowing;
            fastObjectListViewRpdList.FormatRow += fastObjectListViewRpdList_FormatRow;
            // 
            // buttonSelectRpdFiles
            // 
            buttonSelectRpdFiles.Location = new Point(24, 18);
            buttonSelectRpdFiles.Name = "buttonSelectRpdFiles";
            buttonSelectRpdFiles.Size = new Size(143, 23);
            buttonSelectRpdFiles.TabIndex = 3;
            buttonSelectRpdFiles.Text = "Выбор файлов...";
            buttonSelectRpdFiles.UseVisualStyleBackColor = true;
            buttonSelectRpdFiles.Click += buttonSelectRpdFiles_Click;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(label1);
            tabPage2.Controls.Add(fastObjectListViewFosList);
            tabPage2.Controls.Add(textBoxFosDir);
            tabPage2.Controls.Add(buttonSelectFosDir);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1039, 616);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "ФОС";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(21, 24);
            label1.Name = "label1";
            label1.Size = new Size(96, 14);
            label1.TabIndex = 2;
            label1.Text = "Расположение:";
            // 
            // tabPageRpdCheck
            // 
            tabPageRpdCheck.Controls.Add(buttonSaveRpdReport);
            tabPageRpdCheck.Controls.Add(webView2RpdReport);
            tabPageRpdCheck.Location = new Point(4, 24);
            tabPageRpdCheck.Name = "tabPageRpdCheck";
            tabPageRpdCheck.Padding = new Padding(3);
            tabPageRpdCheck.Size = new Size(1039, 616);
            tabPageRpdCheck.TabIndex = 4;
            tabPageRpdCheck.Text = "Проверка РПД";
            tabPageRpdCheck.UseVisualStyleBackColor = true;
            // 
            // buttonSaveRpdReport
            // 
            buttonSaveRpdReport.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            buttonSaveRpdReport.Location = new Point(907, 17);
            buttonSaveRpdReport.Name = "buttonSaveRpdReport";
            buttonSaveRpdReport.Size = new Size(114, 23);
            buttonSaveRpdReport.TabIndex = 6;
            buttonSaveRpdReport.Text = "Сохранить";
            buttonSaveRpdReport.UseVisualStyleBackColor = true;
            buttonSaveRpdReport.Click += buttonSaveRpdReport_Click;
            // 
            // webView2RpdReport
            // 
            webView2RpdReport.AllowExternalDrop = false;
            webView2RpdReport.CreationProperties = null;
            webView2RpdReport.DefaultBackgroundColor = Color.White;
            webView2RpdReport.Dock = DockStyle.Fill;
            webView2RpdReport.Location = new Point(3, 3);
            webView2RpdReport.Name = "webView2RpdReport";
            webView2RpdReport.Size = new Size(1033, 610);
            webView2RpdReport.TabIndex = 5;
            webView2RpdReport.ZoomFactor = 1D;
            // 
            // tabPageRpdGeneration
            // 
            tabPageRpdGeneration.Controls.Add(labelRpdGenStatus);
            tabPageRpdGeneration.Controls.Add(label10);
            tabPageRpdGeneration.Controls.Add(textBoxRpdGenDepartment);
            tabPageRpdGeneration.Controls.Add(label9);
            tabPageRpdGeneration.Controls.Add(textBoxRpdGenProfile);
            tabPageRpdGeneration.Controls.Add(textBoxRpdGenDirectionName);
            tabPageRpdGeneration.Controls.Add(label8);
            tabPageRpdGeneration.Controls.Add(textBoxRpdGenDirectionCode);
            tabPageRpdGeneration.Controls.Add(label7);
            tabPageRpdGeneration.Controls.Add(groupBox2);
            tabPageRpdGeneration.Controls.Add(buttonGenerate);
            tabPageRpdGeneration.Controls.Add(groupBox1);
            tabPageRpdGeneration.Controls.Add(label3);
            tabPageRpdGeneration.Controls.Add(comboBoxRpdGenCurriculumGroups);
            tabPageRpdGeneration.Location = new Point(4, 23);
            tabPageRpdGeneration.Name = "tabPageRpdGeneration";
            tabPageRpdGeneration.Padding = new Padding(3);
            tabPageRpdGeneration.Size = new Size(1039, 617);
            tabPageRpdGeneration.TabIndex = 5;
            tabPageRpdGeneration.Text = "Генерация РПД";
            tabPageRpdGeneration.UseVisualStyleBackColor = true;
            // 
            // labelRpdGenStatus
            // 
            labelRpdGenStatus.AutoEllipsis = true;
            labelRpdGenStatus.Location = new Point(667, 549);
            labelRpdGenStatus.Name = "labelRpdGenStatus";
            labelRpdGenStatus.Size = new Size(361, 52);
            labelRpdGenStatus.TabIndex = 23;
            labelRpdGenStatus.Text = "label11";
            labelRpdGenStatus.Visible = false;
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Location = new Point(704, 49);
            label10.Name = "label10";
            label10.Size = new Size(278, 28);
            label10.TabIndex = 22;
            label10.Text = "<-- Данные поля будут подставлять в шаблон, \r\nздесь их можно переопределить";
            label10.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // textBoxRpdGenDepartment
            // 
            textBoxRpdGenDepartment.Location = new Point(168, 103);
            textBoxRpdGenDepartment.Name = "textBoxRpdGenDepartment";
            textBoxRpdGenDepartment.Size = new Size(481, 22);
            textBoxRpdGenDepartment.TabIndex = 21;
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new Point(101, 106);
            label9.Name = "label9";
            label9.Size = new Size(61, 14);
            label9.TabIndex = 20;
            label9.Text = "Кафедра:";
            // 
            // textBoxRpdGenProfile
            // 
            textBoxRpdGenProfile.Location = new Point(168, 75);
            textBoxRpdGenProfile.Name = "textBoxRpdGenProfile";
            textBoxRpdGenProfile.Size = new Size(481, 22);
            textBoxRpdGenProfile.TabIndex = 19;
            // 
            // textBoxRpdGenDirectionName
            // 
            textBoxRpdGenDirectionName.Location = new Point(274, 47);
            textBoxRpdGenDirectionName.Name = "textBoxRpdGenDirectionName";
            textBoxRpdGenDirectionName.Size = new Size(375, 22);
            textBoxRpdGenDirectionName.TabIndex = 18;
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new Point(99, 78);
            label8.Name = "label8";
            label8.Size = new Size(63, 14);
            label8.TabIndex = 17;
            label8.Text = "Профиль:";
            // 
            // textBoxRpdGenDirectionCode
            // 
            textBoxRpdGenDirectionCode.Location = new Point(168, 47);
            textBoxRpdGenDirectionCode.Name = "textBoxRpdGenDirectionCode";
            textBoxRpdGenDirectionCode.Size = new Size(100, 22);
            textBoxRpdGenDirectionCode.TabIndex = 16;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(76, 50);
            label7.Name = "label7";
            label7.Size = new Size(86, 14);
            label7.TabIndex = 15;
            label7.Text = "Направление:";
            // 
            // groupBox2
            // 
            groupBox2.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            groupBox2.Controls.Add(comboBoxRpdGenTemplates);
            groupBox2.Controls.Add(buttonSelectRpdTargetDir);
            groupBox2.Controls.Add(label4);
            groupBox2.Controls.Add(textBoxRpdGenTargetDir);
            groupBox2.Controls.Add(label6);
            groupBox2.Controls.Add(textBoxRpdGenFileNameTemplate);
            groupBox2.Controls.Add(label5);
            groupBox2.Location = new Point(8, 513);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(622, 103);
            groupBox2.TabIndex = 14;
            groupBox2.TabStop = false;
            groupBox2.Text = "Настройки";
            // 
            // comboBoxRpdGenTemplates
            // 
            comboBoxRpdGenTemplates.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxRpdGenTemplates.FormattingEnabled = true;
            comboBoxRpdGenTemplates.Location = new Point(147, 17);
            comboBoxRpdGenTemplates.Name = "comboBoxRpdGenTemplates";
            comboBoxRpdGenTemplates.Size = new Size(430, 22);
            comboBoxRpdGenTemplates.TabIndex = 14;
            // 
            // buttonSelectRpdTargetDir
            // 
            buttonSelectRpdTargetDir.Location = new Point(583, 73);
            buttonSelectRpdTargetDir.Name = "buttonSelectRpdTargetDir";
            buttonSelectRpdTargetDir.Size = new Size(32, 23);
            buttonSelectRpdTargetDir.TabIndex = 13;
            buttonSelectRpdTargetDir.Text = "...";
            buttonSelectRpdTargetDir.UseVisualStyleBackColor = true;
            buttonSelectRpdTargetDir.Click += buttonSelectRpdTargetDir_Click;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(22, 20);
            label4.Name = "label4";
            label4.Size = new Size(119, 14);
            label4.TabIndex = 5;
            label4.Text = "Шаблон документа:";
            // 
            // textBoxRpdGenTargetDir
            // 
            textBoxRpdGenTargetDir.Location = new Point(147, 73);
            textBoxRpdGenTargetDir.Name = "textBoxRpdGenTargetDir";
            textBoxRpdGenTargetDir.Size = new Size(430, 22);
            textBoxRpdGenTargetDir.TabIndex = 12;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(10, 76);
            label6.Name = "label6";
            label6.Size = new Size(131, 14);
            label6.TabIndex = 11;
            label6.Text = "Целевая директория:";
            // 
            // textBoxRpdGenFileNameTemplate
            // 
            textBoxRpdGenFileNameTemplate.Location = new Point(147, 45);
            textBoxRpdGenFileNameTemplate.Name = "textBoxRpdGenFileNameTemplate";
            textBoxRpdGenFileNameTemplate.Size = new Size(430, 22);
            textBoxRpdGenFileNameTemplate.TabIndex = 10;
            textBoxRpdGenFileNameTemplate.Text = "РПД_{DisciplineName}_2024.docx";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(6, 48);
            label5.Name = "label5";
            label5.Size = new Size(135, 14);
            label5.TabIndex = 9;
            label5.Text = "Шаблон имени файла:";
            // 
            // buttonGenerate
            // 
            buttonGenerate.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonGenerate.Location = new Point(667, 513);
            buttonGenerate.Name = "buttonGenerate";
            buttonGenerate.Size = new Size(164, 23);
            buttonGenerate.TabIndex = 6;
            buttonGenerate.Text = "Сгенерировать РПД...";
            buttonGenerate.UseVisualStyleBackColor = true;
            buttonGenerate.Click += buttonGenerate_Click;
            // 
            // groupBox1
            // 
            groupBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBox1.Controls.Add(fastObjectListViewDisciplineListForGeneration);
            groupBox1.Location = new Point(8, 131);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(1023, 376);
            groupBox1.TabIndex = 4;
            groupBox1.TabStop = false;
            groupBox1.Text = "Выбор дисциплин";
            // 
            // fastObjectListViewDisciplineListForGeneration
            // 
            fastObjectListViewDisciplineListForGeneration.CheckBoxes = true;
            fastObjectListViewDisciplineListForGeneration.Dock = DockStyle.Fill;
            fastObjectListViewDisciplineListForGeneration.FullRowSelect = true;
            fastObjectListViewDisciplineListForGeneration.GridLines = true;
            fastObjectListViewDisciplineListForGeneration.HeaderWordWrap = true;
            fastObjectListViewDisciplineListForGeneration.Location = new Point(3, 18);
            fastObjectListViewDisciplineListForGeneration.Name = "fastObjectListViewDisciplineListForGeneration";
            fastObjectListViewDisciplineListForGeneration.ShowGroups = false;
            fastObjectListViewDisciplineListForGeneration.ShowImagesOnSubItems = true;
            fastObjectListViewDisciplineListForGeneration.Size = new Size(1017, 355);
            fastObjectListViewDisciplineListForGeneration.TabIndex = 4;
            fastObjectListViewDisciplineListForGeneration.UseFilterIndicator = true;
            fastObjectListViewDisciplineListForGeneration.UseFiltering = true;
            fastObjectListViewDisciplineListForGeneration.UseHotItem = true;
            fastObjectListViewDisciplineListForGeneration.UseTranslucentHotItem = true;
            fastObjectListViewDisciplineListForGeneration.UseTranslucentSelection = true;
            fastObjectListViewDisciplineListForGeneration.View = View.Details;
            fastObjectListViewDisciplineListForGeneration.VirtualMode = true;
            fastObjectListViewDisciplineListForGeneration.CellToolTipShowing += fastObjectListViewDisciplineListForGeneration_CellToolTipShowing;
            fastObjectListViewDisciplineListForGeneration.FormatRow += fastObjectListViewDisciplineListForGeneration_FormatRow;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(16, 22);
            label3.Name = "label3";
            label3.Size = new Size(146, 14);
            label3.TabIndex = 1;
            label3.Text = "Группа учебных планов:";
            // 
            // comboBoxRpdGenCurriculumGroups
            // 
            comboBoxRpdGenCurriculumGroups.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxRpdGenCurriculumGroups.FormattingEnabled = true;
            comboBoxRpdGenCurriculumGroups.Location = new Point(168, 19);
            comboBoxRpdGenCurriculumGroups.Name = "comboBoxRpdGenCurriculumGroups";
            comboBoxRpdGenCurriculumGroups.Size = new Size(847, 22);
            comboBoxRpdGenCurriculumGroups.TabIndex = 0;
            comboBoxRpdGenCurriculumGroups.SelectedIndexChanged += comboBoxSelectCurriculum_SelectedIndexChanged;
            // 
            // openFileDialog2
            // 
            openFileDialog2.Filter = "Excel-файлы|*.xlsx|Все файлы|*.*";
            openFileDialog2.Multiselect = true;
            // 
            // openFileDialog3
            // 
            openFileDialog3.Filter = "Word-файлы|*.docx|Все файлы|*.*";
            openFileDialog3.Multiselect = true;
            // 
            // openFileDialog4
            // 
            openFileDialog4.FileName = "openFileDialog2";
            openFileDialog4.Filter = "Excel-файлы|*.xlsx|Все файлы|*.*";
            openFileDialog4.Multiselect = true;
            // 
            // openFileDialogRpdTemplate
            // 
            openFileDialogRpdTemplate.Filter = "Docx-файлы|*.docx|Все файлы|*.*";
            // 
            // FormMain
            // 
            AutoScaleDimensions = new SizeF(7F, 14F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1047, 644);
            Controls.Add(tabControl1);
            Font = new Font("Tahoma", 9F);
            Name = "FormMain";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Менеджер РПД и ФОС";
            Load += FormMain_Load;
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewFosList).EndInit();
            tabControl1.ResumeLayout(false);
            tabPageCompetenceMatrix.ResumeLayout(false);
            tabPageCompetenceMatrix.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)webView21).EndInit();
            tabPageСurriculum.ResumeLayout(false);
            tabPageСurriculum.PerformLayout();
            groupBoxDisciplines.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewDisciplines).EndInit();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewCurricula).EndInit();
            tabPage1.ResumeLayout(false);
            tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewRpdList).EndInit();
            tabPage2.ResumeLayout(false);
            tabPage2.PerformLayout();
            tabPageRpdCheck.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)webView2RpdReport).EndInit();
            tabPageRpdGeneration.ResumeLayout(false);
            tabPageRpdGeneration.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewDisciplineListForGeneration).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private Button buttonSelectFosDir;
        private TextBox textBoxFosDir;
        private BrightIdeasSoftware.FastObjectListView fastObjectListViewFosList;
        private TabControl tabControl1;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private Label label1;
        private TabPage tabPageCompetenceMatrix;
        private Button buttonLoadCompetenceMatrix;
        private Button buttonSelectMatrixFile;
        private Label label2;
        private TextBox textBoxMatrixFileName;
        private OpenFileDialog openFileDialog1;
        private Microsoft.Web.WebView2.WinForms.WebView2 webView21;
        private TabPage tabPageСurriculum;
        private BrightIdeasSoftware.FastObjectListView fastObjectListViewCurricula;
        private Button buttonSelectExcelFiles;
        private OpenFileDialog openFileDialog2;
        private Label labelExcelFileLoading;
        private GroupBox groupBoxDisciplines;
        private BrightIdeasSoftware.FastObjectListView fastObjectListViewDisciplines;
        private Label labelLoadRpd;
        private BrightIdeasSoftware.FastObjectListView fastObjectListViewRpdList;
        private Button buttonSelectRpdFiles;
        private OpenFileDialog openFileDialog3;
        private OpenFileDialog openFileDialog4;
        private Button buttonRpdCheck;
        private TabPage tabPageRpdCheck;
        private Microsoft.Web.WebView2.WinForms.WebView2 webView2RpdReport;
        private Button buttonSaveRpdReport;
        private TabPage tabPageRpdGeneration;
        private Label label3;
        private ComboBox comboBoxRpdGenCurriculumGroups;
        private GroupBox groupBox1;
        private BrightIdeasSoftware.FastObjectListView fastObjectListViewDisciplineListForGeneration;
        private Button buttonSelectRpdTemplate;
        private TextBox textBoxRpdTemplate;
        private Button buttonGenerate;
        private Label label4;
        private OpenFileDialog openFileDialogRpdTemplate;
        private TextBox textBoxRpdGenFileNameTemplate;
        private Label label5;
        private Button buttonSelectRpdTargetDir;
        private TextBox textBoxRpdGenTargetDir;
        private Label label6;
        private GroupBox groupBox2;
        private FolderBrowserDialog folderBrowserDialogRpdTargetDir;
        private TextBox textBoxRpdGenDirectionCode;
        private Label label7;
        private TextBox textBoxRpdGenDepartment;
        private Label label9;
        private TextBox textBoxRpdGenProfile;
        private TextBox textBoxRpdGenDirectionName;
        private Label label8;
        private ComboBox comboBoxRpdGenTemplates;
        private Label label10;
        private Label labelRpdGenStatus;
    }
}
