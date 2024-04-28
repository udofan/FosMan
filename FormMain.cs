namespace FosMan {
    public partial class FormMain : Form {
        CompetenceMatrix m_matrix = null;
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
                MessageBox.Show("���� ������� ����������� �� �����.", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                loadMatrix = false;
            }

            if (loadMatrix && m_matrix != null) {
                if (MessageBox.Show("������� ��� ���������.\r\n��������?", "��������", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation) != DialogResult.Yes) {
                    loadMatrix = false;
                }
            }

            if (loadMatrix) {
                m_matrix = CompetenceMatrix.LoadFromFile(textBoxMatrixFileName.Text, out var errors);

                if (m_matrix != null) {
                    m_matrixFileName = textBoxMatrixFileName.Text;

                    ShowMatrix(errors);
                    MessageBox.Show("������� ����������� ������� ���������.", "��������", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else {
                    var errorList = string.Join("\r\n", errors);
                    MessageBox.Show($"���������� ������:\r\n{errorList}", "�������� �������", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
        }

        /// <summary>
        /// ������� ������� ����������� � WebView2
        /// </summary>
        /// <param name="errors"></param>
        async void ShowMatrix(List<string> errors) {
            var html = "<html><body>";
            var tdStyle = " style='border: 1px solid; vertical-align: top'";

            //���� ���� ������
            if (errors.Any()) {
                html += "<div style='color: red'><b>������:</b></div>";
                html += string.Join("", errors.Select(e => $"<div style='color: red'>{e}</div>"));
            }

            //������������ �������
            html += "<table style='border: 1px solid'>";
            html += $"<tr><th {tdStyle}><b>��� � ������������ �����������</b></th>" +
                        $"<th {tdStyle}><b>���� � ���������� ���������� �����������\r\n</b></tdh>" +
                        $"<th {tdStyle}><b>���� � ���������� ��������</b></td></th></tr>";
            foreach (var item in m_matrix.Items) {
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
    }
}
