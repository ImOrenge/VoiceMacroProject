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
        private Label statusLabel;

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
            this.ClientSize = new System.Drawing.Size(700, 550);
            this.Name = "PresetManagerForm";
            this.Text = "프리셋 관리";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.Padding = new Padding(15);
            
            // 제목 레이블 추가
            Label titleLabel = new Label();
            titleLabel.Text = "프리셋 관리";
            titleLabel.Font = new Font(this.Font.FontFamily, 12, FontStyle.Bold);
            titleLabel.Location = new Point(15, 15);
            titleLabel.Size = new Size(350, 25);
            titleLabel.Padding = new Padding(5, 0, 0, 0);
            this.Controls.Add(titleLabel);
            
            // 설명 레이블 추가
            Label descLabel = new Label();
            descLabel.Text = "사용 가능한 프리셋을 관리하고 불러옵니다.";
            descLabel.Location = new Point(20, 40);
            descLabel.Size = new Size(550, 20);
            descLabel.ForeColor = Color.DarkGray;
            this.Controls.Add(descLabel);

            // 프리셋 목록 그룹
            GroupBox presetListGroup = new GroupBox();
            presetListGroup.Text = "사용 가능한 프리셋";
            presetListGroup.Location = new Point(20, 70);
            presetListGroup.Size = new Size(660, 350);
            presetListGroup.Padding = new Padding(10);
            presetListGroup.BackColor = Color.FromArgb(245, 245, 245);
            this.Controls.Add(presetListGroup);

            // 프리셋 목록
            lstPresets = new ListView();
            lstPresets.Location = new Point(15, 25);
            lstPresets.Size = new Size(630, 310);
            lstPresets.View = View.Details;
            lstPresets.FullRowSelect = true;
            lstPresets.MultiSelect = false;
            lstPresets.HideSelection = false;
            lstPresets.GridLines = true;
            lstPresets.BackColor = Color.White;
            lstPresets.BorderStyle = BorderStyle.Fixed3D;
            lstPresets.Columns.Add("이름", 180);
            lstPresets.Columns.Add("매크로 수", 80);
            lstPresets.Columns.Add("수정 날짜", 180);
            lstPresets.Columns.Add("경로", 170);
            lstPresets.SelectedIndexChanged += LstPresets_SelectedIndexChanged;
            // 대체 행 색상 설정을 위한 이벤트 추가
            lstPresets.DrawItem += (s, e) => {
                if (e.ItemIndex % 2 == 1)
                {
                    e.Item.BackColor = Color.FromArgb(248, 248, 248);
                }
            };
            presetListGroup.Controls.Add(lstPresets);

            // 버튼 그룹 영역
            Panel buttonPanel = new Panel();
            buttonPanel.Location = new Point(20, 430);
            buttonPanel.Size = new Size(660, 85);
            this.Controls.Add(buttonPanel);

            // 프리셋 관리 그룹
            GroupBox manageGroup = new GroupBox();
            manageGroup.Text = "프리셋 관리";
            manageGroup.Location = new Point(0, 0);
            manageGroup.Size = new Size(380, 80);
            buttonPanel.Controls.Add(manageGroup);

            // 프리셋 관리 버튼들 - 크기 확대 및 간격 조정
            btnNew = new Button();
            btnNew.Text = "새 프리셋 저장";
            btnNew.Location = new Point(15, 30);
            btnNew.Size = new Size(110, 35);
            btnNew.BackColor = Color.LightBlue;
            btnNew.FlatStyle = FlatStyle.Flat;
            btnNew.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
            btnNew.Click += BtnNew_Click;
            manageGroup.Controls.Add(btnNew);

            btnLoad = new Button();
            btnLoad.Text = "불러오기";
            btnLoad.Location = new Point(135, 30);
            btnLoad.Size = new Size(110, 35);
            btnLoad.FlatStyle = FlatStyle.Flat;
            btnLoad.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
            btnLoad.Enabled = false;
            btnLoad.Click += BtnLoad_Click;
            manageGroup.Controls.Add(btnLoad);

            btnDelete = new Button();
            btnDelete.Text = "삭제";
            btnDelete.Location = new Point(255, 30);
            btnDelete.Size = new Size(110, 35);
            btnDelete.FlatStyle = FlatStyle.Flat;
            btnDelete.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
            btnDelete.Enabled = false;
            btnDelete.Click += BtnDelete_Click;
            manageGroup.Controls.Add(btnDelete);

            // 가져오기/내보내기 그룹
            GroupBox importExportGroup = new GroupBox();
            importExportGroup.Text = "가져오기/내보내기";
            importExportGroup.Location = new Point(390, 0);
            importExportGroup.Size = new Size(270, 80);
            buttonPanel.Controls.Add(importExportGroup);

            btnImport = new Button();
            btnImport.Text = "선택 가져오기";
            btnImport.Location = new Point(15, 30);
            btnImport.Size = new Size(110, 35);
            btnImport.FlatStyle = FlatStyle.Flat;
            btnImport.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
            btnImport.Enabled = false;
            btnImport.Click += BtnImport_Click;
            importExportGroup.Controls.Add(btnImport);

            btnExport = new Button();
            btnExport.Text = "외부에서 가져오기";
            btnExport.Location = new Point(135, 30);
            btnExport.Size = new Size(120, 35);
            btnExport.FlatStyle = FlatStyle.Flat;
            btnExport.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
            btnExport.Click += BtnExport_Click;
            importExportGroup.Controls.Add(btnExport);

            // 닫기 버튼
            btnClose = new Button();
            btnClose.Text = "닫기";
            btnClose.Location = new Point(590, 520);
            btnClose.Size = new Size(90, 35);
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
            btnClose.Click += (sender, e) => this.Close();
            this.Controls.Add(btnClose);

            // 상태 레이블 추가
            statusLabel = new Label();
            statusLabel.Text = "프리셋을 선택하거나 새로 저장하세요.";
            statusLabel.Location = new Point(20, 520);
            statusLabel.Size = new Size(560, 20);
            statusLabel.ForeColor = Color.Gray;
            statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.Controls.Add(statusLabel);

            this.ResumeLayout(false);
        }

        private void LstPresets_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasSelection = lstPresets.SelectedItems.Count > 0;
            btnLoad.Enabled = hasSelection;
            btnImport.Enabled = hasSelection;
            btnDelete.Enabled = hasSelection;
            
            // 선택된 항목에 대한 상태 업데이트
            if (hasSelection)
            {
                var preset = (PresetInfo)lstPresets.SelectedItems[0].Tag;
                statusLabel.Text = $"선택된 프리셋: {preset.Name} (매크로 {preset.MacroCount}개)";
            }
            else
            {
                statusLabel.Text = "프리셋을 선택하거나 새로 저장하세요.";
            }
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
            
            // 상태 업데이트
            statusLabel.Text = $"총 {availablePresets.Count}개의 프리셋이 있습니다.";
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
                            statusLabel.Text = $"새 프리셋 '{presetName}'이(가) 저장되었습니다.";
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
                        statusLabel.Text = $"프리셋 '{preset.Name}'을(를) 불러왔습니다.";
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
                        statusLabel.Text = $"프리셋 '{preset.Name}'을(를) 가져왔습니다.";
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
                        statusLabel.Text = $"프리셋 '{preset.Name}'이(가) 삭제되었습니다.";
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
                            statusLabel.Text = "외부 프리셋을 성공적으로 가져왔습니다.";
                            this.DialogResult = DialogResult.OK; // 성공 시 폼 닫기
                            this.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"외부 파일 가져오기 오류: {ex.Message}", "오류", 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        statusLabel.Text = "외부 파일 가져오기 실패";
                    }
                }
            }
        }

        private void MacroService_StatusChanged(object sender, string message)
        {
            // 상태 레이블에 메시지 표시
            if (statusLabel.InvokeRequired)
            {
                statusLabel.Invoke(new Action(() => statusLabel.Text = message));
            }
            else
            {
                statusLabel.Text = message;
            }
            
            // 중요한 메시지는 대화상자로도 표시
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
            this.Size = new Size(400, 180);
            this.BackColor = Color.FromArgb(240, 240, 240);

            lblPrompt = new Label();
            lblPrompt.Text = prompt;
            lblPrompt.Location = new Point(20, 20);
            lblPrompt.Size = new Size(360, 20);
            lblPrompt.Font = new Font(this.Font, FontStyle.Regular);
            this.Controls.Add(lblPrompt);

            txtInput = new TextBox();
            txtInput.Location = new Point(20, 50);
            txtInput.Size = new Size(360, 25);
            txtInput.Font = new Font(this.Font.FontFamily, 10);
            this.Controls.Add(txtInput);

            btnOK = new Button();
            btnOK.Text = "확인";
            btnOK.DialogResult = DialogResult.OK;
            btnOK.Location = new Point(210, 100);
            btnOK.Size = new Size(80, 30);
            btnOK.FlatStyle = FlatStyle.Flat;
            btnOK.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
            btnOK.BackColor = Color.LightBlue;
            this.Controls.Add(btnOK);
            this.AcceptButton = btnOK;

            btnCancel = new Button();
            btnCancel.Text = "취소";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(300, 100);
            btnCancel.Size = new Size(80, 30);
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
            this.Controls.Add(btnCancel);
            this.CancelButton = btnCancel;
        }
    }
} 