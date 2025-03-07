using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using VoiceMacro.Services;

namespace VoiceMacro
{
    public partial class PresetManagerForm : Form
    {
        private readonly MacroService macroService;
        private List<PresetInfo> availablePresets;
        private ListView lstPresets;
        private Button btnNew;
        private Button btnLoad;
        private Button btnImport;
        private Button btnDelete;
        private Button btnClose;
        private Button btnExport;

        public PresetManagerForm(MacroService macroService)
        {
            this.macroService = macroService;
            this.macroService.StatusChanged += MacroService_StatusChanged;

            InitializeComponent();
            LoadPresets();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // 폼 설정
            this.ClientSize = new System.Drawing.Size(500, 400);
            this.Name = "PresetManagerForm";
            this.Text = "프리셋 관리";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 프리셋 목록
            lstPresets = new ListView();
            lstPresets.Location = new Point(20, 20);
            lstPresets.Size = new Size(460, 300);
            lstPresets.View = View.Details;
            lstPresets.FullRowSelect = true;
            lstPresets.MultiSelect = false;
            lstPresets.HideSelection = false;
            lstPresets.Columns.Add("이름", 150);
            lstPresets.Columns.Add("매크로 수", 80);
            lstPresets.Columns.Add("수정 날짜", 150);
            lstPresets.Columns.Add("경로", 150);
            lstPresets.SelectedIndexChanged += LstPresets_SelectedIndexChanged;
            this.Controls.Add(lstPresets);

            // 버튼 영역
            btnNew = new Button();
            btnNew.Text = "새 프리셋 저장";
            btnNew.Location = new Point(20, 330);
            btnNew.Size = new Size(100, 30);
            btnNew.Click += BtnNew_Click;
            this.Controls.Add(btnNew);

            btnLoad = new Button();
            btnLoad.Text = "불러오기";
            btnLoad.Location = new Point(130, 330);
            btnLoad.Size = new Size(80, 30);
            btnLoad.Enabled = false;
            btnLoad.Click += BtnLoad_Click;
            this.Controls.Add(btnLoad);

            btnImport = new Button();
            btnImport.Text = "가져오기";
            btnImport.Location = new Point(220, 330);
            btnImport.Size = new Size(80, 30);
            btnImport.Enabled = false;
            btnImport.Click += BtnImport_Click;
            this.Controls.Add(btnImport);

            btnDelete = new Button();
            btnDelete.Text = "삭제";
            btnDelete.Location = new Point(310, 330);
            btnDelete.Size = new Size(70, 30);
            btnDelete.Enabled = false;
            btnDelete.Click += BtnDelete_Click;
            this.Controls.Add(btnDelete);

            btnClose = new Button();
            btnClose.Text = "닫기";
            btnClose.Location = new Point(390, 330);
            btnClose.Size = new Size(70, 30);
            btnClose.Click += (sender, e) => this.Close();
            this.Controls.Add(btnClose);

            // 외부 프리셋 가져오기 버튼
            btnExport = new Button();
            btnExport.Text = "외부 파일에서 가져오기";
            btnExport.Location = new Point(20, 370);
            btnExport.Size = new Size(150, 30);
            btnExport.Click += BtnExport_Click;
            this.Controls.Add(btnExport);

            this.ResumeLayout(false);
        }

        private void LstPresets_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasSelection = lstPresets.SelectedItems.Count > 0;
            btnLoad.Enabled = hasSelection;
            btnImport.Enabled = hasSelection;
            btnDelete.Enabled = hasSelection;
        }

        private void LoadPresets()
        {
            lstPresets.Items.Clear();
            availablePresets = macroService.GetAvailablePresets();

            foreach (var preset in availablePresets)
            {
                var item = new ListViewItem(preset.Name);
                item.SubItems.Add(preset.MacroCount.ToString());
                item.SubItems.Add(preset.LastModified.ToString("yyyy-MM-dd HH:mm:ss"));
                item.SubItems.Add(preset.FilePath);
                item.Tag = preset; // 프리셋 정보 저장
                lstPresets.Items.Add(item);
            }
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            using (var inputDialog = new InputDialog("새 프리셋 저장", "프리셋 이름을 입력하세요:"))
            {
                if (inputDialog.ShowDialog() == DialogResult.OK)
                {
                    string presetName = inputDialog.InputText.Trim();
                    if (!string.IsNullOrEmpty(presetName))
                    {
                        if (macroService.SavePreset(presetName))
                        {
                            LoadPresets(); // 목록 새로고침
                        }
                    }
                }
            }
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (lstPresets.SelectedItems.Count > 0)
            {
                var preset = (PresetInfo)lstPresets.SelectedItems[0].Tag;
                
                DialogResult result = MessageBox.Show(
                    $"'{preset.Name}' 프리셋을 불러오시겠습니까?\n현재 매크로 설정이 모두 교체됩니다.",
                    "프리셋 불러오기",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    if (macroService.LoadPreset(preset.FilePath))
                    {
                        this.DialogResult = DialogResult.OK; // 성공 시 폼 닫기
                        this.Close();
                    }
                }
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            if (lstPresets.SelectedItems.Count > 0)
            {
                var preset = (PresetInfo)lstPresets.SelectedItems[0].Tag;

                DialogResult result = MessageBox.Show(
                    $"'{preset.Name}' 프리셋을 현재 매크로 설정에 추가하시겠습니까?",
                    "프리셋 가져오기",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    if (macroService.ImportPreset(preset.FilePath))
                    {
                        this.DialogResult = DialogResult.OK; // 성공 시 폼 닫기
                        this.Close();
                    }
                }
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (lstPresets.SelectedItems.Count > 0)
            {
                var preset = (PresetInfo)lstPresets.SelectedItems[0].Tag;

                DialogResult result = MessageBox.Show(
                    $"'{preset.Name}' 프리셋을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.",
                    "프리셋 삭제",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    if (macroService.DeletePreset(preset.FilePath))
                    {
                        LoadPresets(); // 목록 새로고침
                    }
                }
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "JSON 파일 (*.json)|*.json";
                openFileDialog.Title = "외부 프리셋 파일 선택";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (macroService.ImportPreset(openFileDialog.FileName))
                        {
                            this.DialogResult = DialogResult.OK; // 성공 시 폼 닫기
                            this.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"외부 파일 가져오기 오류: {ex.Message}", "오류", 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void MacroService_StatusChanged(object sender, string message)
        {
            // 메시지 표시 (상태 표시줄이 있을 경우)
            MessageBox.Show(message, "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>
    /// 텍스트 입력을 받기 위한 간단한 대화상자
    /// </summary>
    public class InputDialog : Form
    {
        private TextBox txtInput;
        private Button btnOK;
        private Button btnCancel;
        private Label lblPrompt;

        public string InputText
        {
            get { return txtInput.Text; }
        }

        public InputDialog(string title, string prompt)
        {
            this.Text = title;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Size = new Size(350, 150);

            lblPrompt = new Label();
            lblPrompt.Text = prompt;
            lblPrompt.Location = new Point(10, 20);
            lblPrompt.Size = new Size(320, 20);
            this.Controls.Add(lblPrompt);

            txtInput = new TextBox();
            txtInput.Location = new Point(10, 50);
            txtInput.Size = new Size(320, 20);
            this.Controls.Add(txtInput);

            btnOK = new Button();
            btnOK.Text = "확인";
            btnOK.DialogResult = DialogResult.OK;
            btnOK.Location = new Point(165, 80);
            btnOK.Size = new Size(75, 23);
            this.Controls.Add(btnOK);
            this.AcceptButton = btnOK;

            btnCancel = new Button();
            btnCancel.Text = "취소";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(255, 80);
            btnCancel.Size = new Size(75, 23);
            this.Controls.Add(btnCancel);
            this.CancelButton = btnCancel;
        }
    }
} 