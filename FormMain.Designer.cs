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
        private void InitializeComponent()
        {
            buttonSelectFosDir = new Button();
            textBoxFosDir = new TextBox();
            fastObjectListViewFosList = new BrightIdeasSoftware.FastObjectListView();
            tabControl1 = new TabControl();
            tabPage1 = new TabPage();
            tabPage2 = new TabPage();
            label1 = new Label();
            tabPage3 = new TabPage();
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewFosList).BeginInit();
            tabControl1.SuspendLayout();
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
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage3);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1047, 644);
            tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            tabPage1.Location = new Point(4, 23);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(661, 492);
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
            tabPage2.Location = new Point(4, 23);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1039, 617);
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
            // tabPage3
            // 
            tabPage3.Location = new Point(4, 23);
            tabPage3.Name = "tabPage3";
            tabPage3.Padding = new Padding(3);
            tabPage3.Size = new Size(1039, 617);
            tabPage3.TabIndex = 2;
            tabPage3.Text = "Матрица компетенций";
            tabPage3.UseVisualStyleBackColor = true;
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
            ((System.ComponentModel.ISupportInitialize)fastObjectListViewFosList).EndInit();
            tabControl1.ResumeLayout(false);
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
        private TabPage tabPage3;
    }
}
