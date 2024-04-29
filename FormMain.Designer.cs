﻿namespace FosMan {
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
            tabPage2 = new TabPage();
            label1 = new Label();
            openFileDialog1 = new OpenFileDialog();
            openFileDialog2 = new OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewFosList).BeginInit();
            tabControl1.SuspendLayout();
            tabPageCompetenceMatrix.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)webView21).BeginInit();
            tabPageСurriculum.SuspendLayout();
            groupBoxDisciplines.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewDisciplines).BeginInit();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewCurricula).BeginInit();
            tabPage2.SuspendLayout();
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
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1047, 644);
            tabControl1.TabIndex = 1;
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
            webView21.Location = new Point(20, 59);
            webView21.Name = "webView21";
            webView21.Size = new Size(1001, 546);
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
            textBoxMatrixFileName.Text = "c:\\FosMan\\Матрицы_компетенций\\1. Индикаторы и компетенции_37.03.01 _Пс  3++.docx";
            // 
            // tabPageСurriculum
            // 
            tabPageСurriculum.Controls.Add(groupBoxDisciplines);
            tabPageСurriculum.Controls.Add(labelExcelFileLoading);
            tabPageСurriculum.Controls.Add(fastObjectListViewCurricula);
            tabPageСurriculum.Controls.Add(buttonSelectExcelFiles);
            tabPageСurriculum.Location = new Point(4, 23);
            tabPageСurriculum.Name = "tabPageСurriculum";
            tabPageСurriculum.Padding = new Padding(3);
            tabPageСurriculum.Size = new Size(1039, 617);
            tabPageСurriculum.TabIndex = 3;
            tabPageСurriculum.Text = "Учебные планы";
            tabPageСurriculum.UseVisualStyleBackColor = true;
            // 
            // groupBoxDisciplines
            // 
            groupBoxDisciplines.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBoxDisciplines.Controls.Add(fastObjectListViewDisciplines);
            groupBoxDisciplines.Location = new Point(8, 358);
            groupBoxDisciplines.Name = "groupBoxDisciplines";
            groupBoxDisciplines.Size = new Size(1023, 257);
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
            fastObjectListViewDisciplines.Size = new Size(1017, 236);
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
            fastObjectListViewCurricula.Size = new Size(1023, 295);
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
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(1039, 616);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "РПД";
            tabPage1.UseVisualStyleBackColor = true;
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
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // openFileDialog2
            // 
            openFileDialog2.FileName = "openFileDialog2";
            openFileDialog2.Filter = "Excel-файлы|*.xlsx|Все файлы|*.*";
            openFileDialog2.Multiselect = true;
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
            tabPage2.ResumeLayout(false);
            tabPage2.PerformLayout();
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
    }
}
