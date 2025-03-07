using System;
using System.Windows.Forms;
using VoiceMacro.Services;

namespace VoiceMacro
{
    public partial class SettingsForm : Form
    {
        private readonly VoiceRecognitionService voiceRecognitionService;
        private readonly AppSettings settings;

        public SettingsForm(VoiceRecognitionService voiceRecognitionService)
        {
            this.voiceRecognitionService = voiceRecognitionService;
            settings = AppSettings.Load();
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // 폼 설정
            this.ClientSize = new System.Drawing.Size(500, 400);
            this.Name = "SettingsForm";
            this.Text = "음성 매크로 설정";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // API 키 설정
            Label lblApiKey = new Label();
            lblApiKey.Text = "OpenAI API 키:";
            lblApiKey.Location = new System.Drawing.Point(20, 20);
            lblApiKey.Size = new System.Drawing.Size(150, 20);
            this.Controls.Add(lblApiKey);

            TextBox txtApiKey = new TextBox();
            txtApiKey.Location = new System.Drawing.Point(180, 20);
            txtApiKey.Size = new System.Drawing.Size(280, 20);
            txtApiKey.PasswordChar = '*';
            txtApiKey.Text = settings.OpenAIApiKey ?? "";
            this.Controls.Add(txtApiKey);

            // API 사용 여부
            CheckBox chkUseApi = new CheckBox();
            chkUseApi.Text = "OpenAI API 사용 (정확도 향상)";
            chkUseApi.Location = new System.Drawing.Point(180, 50);
            chkUseApi.Size = new System.Drawing.Size(250, 20);
            chkUseApi.Checked = settings.UseOpenAIApi;
            chkUseApi.Enabled = !string.IsNullOrEmpty(settings.OpenAIApiKey);
            txtApiKey.TextChanged += (s, e) => 
            {
                chkUseApi.Enabled = !string.IsNullOrEmpty(txtApiKey.Text);
                if (string.IsNullOrEmpty(txtApiKey.Text))
                {
                    chkUseApi.Checked = false;
                }
            };
            this.Controls.Add(chkUseApi);

            // 언어 설정
            Label lblLanguage = new Label();
            lblLanguage.Text = "인식 언어:";
            lblLanguage.Location = new System.Drawing.Point(20, 80);
            lblLanguage.Size = new System.Drawing.Size(150, 20);
            this.Controls.Add(lblLanguage);

            ComboBox cmbLanguage = new ComboBox();
            cmbLanguage.Location = new System.Drawing.Point(180, 80);
            cmbLanguage.Size = new System.Drawing.Size(150, 20);
            cmbLanguage.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLanguage.Items.AddRange(new string[] { "ko", "en", "ja", "zh", "fr", "de", "es", "ru" });
            cmbLanguage.SelectedItem = settings.WhisperLanguage;
            if (cmbLanguage.SelectedIndex == -1 && cmbLanguage.Items.Count > 0)
            {
                cmbLanguage.SelectedIndex = 0;
            }
            this.Controls.Add(cmbLanguage);

            // 녹음 감도 설정
            Label lblThreshold = new Label();
            lblThreshold.Text = "음성 감지 임계값 (dB):";
            lblThreshold.Location = new System.Drawing.Point(20, 120);
            lblThreshold.Size = new System.Drawing.Size(150, 20);
            this.Controls.Add(lblThreshold);

            TrackBar trkThreshold = new TrackBar();
            trkThreshold.Location = new System.Drawing.Point(180, 110);
            trkThreshold.Size = new System.Drawing.Size(200, 45);
            trkThreshold.Minimum = -60;
            trkThreshold.Maximum = -10;
            trkThreshold.Value = settings.RecordingThresholdDb;
            trkThreshold.TickFrequency = 5;
            this.Controls.Add(trkThreshold);

            Label lblThresholdValue = new Label();
            lblThresholdValue.Text = trkThreshold.Value.ToString() + " dB";
            lblThresholdValue.Location = new System.Drawing.Point(390, 120);
            lblThresholdValue.Size = new System.Drawing.Size(70, 20);
            trkThreshold.ValueChanged += (s, e) => lblThresholdValue.Text = trkThreshold.Value.ToString() + " dB";
            this.Controls.Add(lblThresholdValue);

            // 최소 녹음 시간
            Label lblMinDuration = new Label();
            lblMinDuration.Text = "최소 녹음 시간 (ms):";
            lblMinDuration.Location = new System.Drawing.Point(20, 160);
            lblMinDuration.Size = new System.Drawing.Size(150, 20);
            this.Controls.Add(lblMinDuration);

            NumericUpDown numMinDuration = new NumericUpDown();
            numMinDuration.Location = new System.Drawing.Point(180, 160);
            numMinDuration.Size = new System.Drawing.Size(100, 20);
            numMinDuration.Minimum = 500;
            numMinDuration.Maximum = 5000;
            numMinDuration.Increment = 100;
            numMinDuration.Value = settings.MinimumRecordingDuration;
            this.Controls.Add(numMinDuration);

            // 최대 녹음 시간
            Label lblMaxDuration = new Label();
            lblMaxDuration.Text = "최대 녹음 시간 (ms):";
            lblMaxDuration.Location = new System.Drawing.Point(20, 190);
            lblMaxDuration.Size = new System.Drawing.Size(150, 20);
            this.Controls.Add(lblMaxDuration);

            NumericUpDown numMaxDuration = new NumericUpDown();
            numMaxDuration.Location = new System.Drawing.Point(180, 190);
            numMaxDuration.Size = new System.Drawing.Size(100, 20);
            numMaxDuration.Minimum = 5000;
            numMaxDuration.Maximum = 60000;
            numMaxDuration.Increment = 1000;
            numMaxDuration.Value = settings.MaximumRecordingDuration;
            this.Controls.Add(numMaxDuration);

            // 침묵 시간
            Label lblSilenceTimeout = new Label();
            lblSilenceTimeout.Text = "침묵 감지 시간 (ms):";
            lblSilenceTimeout.Location = new System.Drawing.Point(20, 220);
            lblSilenceTimeout.Size = new System.Drawing.Size(150, 20);
            this.Controls.Add(lblSilenceTimeout);

            NumericUpDown numSilenceTimeout = new NumericUpDown();
            numSilenceTimeout.Location = new System.Drawing.Point(180, 220);
            numSilenceTimeout.Size = new System.Drawing.Size(100, 20);
            numSilenceTimeout.Minimum = 500;
            numSilenceTimeout.Maximum = 5000;
            numSilenceTimeout.Increment = 100;
            numSilenceTimeout.Value = settings.SilenceTimeout;
            this.Controls.Add(numSilenceTimeout);

            // 안내 레이블
            Label lblInfo = new Label();
            lblInfo.Text = "OpenAI API를 사용하면 로컬 모델보다 더 정확한 음성 인식이 가능합니다.\n" +
                         "그러나 API 사용 비용이 발생할 수 있습니다.";
            lblInfo.Location = new System.Drawing.Point(20, 270);
            lblInfo.Size = new System.Drawing.Size(450, 40);
            lblInfo.AutoSize = true;
            this.Controls.Add(lblInfo);

            // 버튼 영역
            Button btnSave = new Button();
            btnSave.Text = "저장";
            btnSave.Location = new System.Drawing.Point(200, 330);
            btnSave.Size = new System.Drawing.Size(100, 30);
            btnSave.Click += (sender, e) =>
            {
                // 설정 저장
                settings.OpenAIApiKey = txtApiKey.Text.Trim();
                settings.UseOpenAIApi = chkUseApi.Checked;
                settings.WhisperLanguage = cmbLanguage.SelectedItem?.ToString() ?? "ko";
                settings.RecordingThresholdDb = trkThreshold.Value;
                settings.MinimumRecordingDuration = (int)numMinDuration.Value;
                settings.MaximumRecordingDuration = (int)numMaxDuration.Value;
                settings.SilenceTimeout = (int)numSilenceTimeout.Value;
                
                settings.Save();
                
                // VoiceRecognitionService에 설정 업데이트
                voiceRecognitionService.UpdateSettings(settings);
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            this.Controls.Add(btnSave);

            Button btnCancel = new Button();
            btnCancel.Text = "취소";
            btnCancel.Location = new System.Drawing.Point(310, 330);
            btnCancel.Size = new System.Drawing.Size(100, 30);
            btnCancel.Click += (sender, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };
            this.Controls.Add(btnCancel);

            this.ResumeLayout(false);
        }
    }
} 