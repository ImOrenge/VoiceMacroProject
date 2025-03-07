using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VoiceMacro.Services;

namespace VoiceMacro
{
    public partial class AddMacroForm : Form
    {
        public string Keyword { get; private set; }
        public string KeyAction { get; private set; }

        private readonly AppSettings settings;
        private readonly AudioRecordingService audioRecorder;
        private readonly OpenAIService openAIService;
        private readonly VoiceRecognitionService voiceRecognitionService;
        private TextBox txtKeyword;
        private TextBox txtAction;
        private Button btnRecord;
        private CancellationTokenSource cancellationTokenSource;
        private ProgressBar progressRecording;
        private Label lblRecordingStatus;
        private bool isRecording = false;

        public AddMacroForm(VoiceRecognitionService voiceRecognitionService)
        {
            this.voiceRecognitionService = voiceRecognitionService;
            settings = AppSettings.Load();
            
            // 오디오 녹음 서비스 초기화
            audioRecorder = new AudioRecordingService(settings);
            audioRecorder.RecordingStatusChanged += AudioRecorder_RecordingStatusChanged;
            audioRecorder.AudioLevelChanged += AudioRecorder_AudioLevelChanged;
            
            // OpenAI 서비스 (API 키가 있는 경우)
            if (!string.IsNullOrEmpty(settings.OpenAIApiKey))
            {
                openAIService = new OpenAIService(settings.OpenAIApiKey);
            }
            
            InitializeComponent();
        }

        private void AudioRecorder_RecordingStatusChanged(object sender, string status)
        {
            // UI 스레드에서 처리
            if (lblRecordingStatus.InvokeRequired)
            {
                lblRecordingStatus.Invoke(new Action(() => lblRecordingStatus.Text = status));
            }
            else
            {
                lblRecordingStatus.Text = status;
            }
        }

        private void AudioRecorder_AudioLevelChanged(object sender, float level)
        {
            // UI 스레드에서 처리
            if (progressRecording.InvokeRequired)
            {
                progressRecording.Invoke(new Action(() => 
                {
                    // dB 값(-60 ~ 0)을 0-100 범위로 변환
                    int percentage = (int)Math.Min(100, Math.Max(0, (level + 60) * 100 / 60));
                    progressRecording.Value = percentage;
                }));
            }
            else
            {
                int percentage = (int)Math.Min(100, Math.Max(0, (level + 60) * 100 / 60));
                progressRecording.Value = percentage;
            }
        }

        private async void BtnRecord_Click(object sender, EventArgs e)
        {
            if (isRecording)
            {
                // 녹음 중지
                cancellationTokenSource?.Cancel();
                btnRecord.Text = "음성 녹음";
                isRecording = false;
                return;
            }

            try
            {
                // 마이크 감지 확인
                if (!audioRecorder.HasMicrophone())
                {
                    MessageBox.Show("연결된 마이크가 없습니다. 마이크를 연결한 후 다시 시도하세요.", 
                        "마이크 감지 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    lblRecordingStatus.Text = "마이크 감지 실패";
                    return;
                }

                // 녹음 시작
                btnRecord.Text = "중지";
                isRecording = true;
                cancellationTokenSource = new CancellationTokenSource();
                
                // 음성 녹음 및 인식 시작
                byte[] audioData = await audioRecorder.RecordSpeechAsync(cancellationTokenSource.Token);
                
                if (audioData.Length > 0)
                {
                    lblRecordingStatus.Text = "음성 인식 중...";
                    
                    // 음성 인식 (OpenAI API 또는 로컬)
                    string recognizedText;
                    
                    if (settings.UseOpenAIApi && openAIService != null)
                    {
                        recognizedText = await openAIService.TranscribeAudioAsync(
                            audioData, settings.WhisperLanguage, CancellationToken.None);
                    }
                    else
                    {
                        // VoiceRecognitionService의 로컬 Whisper 프로세서를 활용
                        recognizedText = await voiceRecognitionService.RecognizeAudioAsync(
                            audioData, settings.WhisperLanguage);
                    }
                    
                    if (!string.IsNullOrWhiteSpace(recognizedText))
                    {
                        txtKeyword.Text = recognizedText.Trim();
                        lblRecordingStatus.Text = "인식 완료";
                    }
                    else
                    {
                        lblRecordingStatus.Text = "인식 실패";
                    }
                }
                else
                {
                    // 오디오 데이터가 없는 경우 (마이크 감지 실패 등)
                    lblRecordingStatus.Text = "녹음 실패";
                }
                
                btnRecord.Text = "음성 녹음";
                isRecording = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"녹음 중 오류가 발생했습니다: {ex.Message}", "오류", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                btnRecord.Text = "음성 녹음";
                isRecording = false;
                lblRecordingStatus.Text = "오류 발생";
            }
            finally
            {
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // 폼 설정
            this.ClientSize = new System.Drawing.Size(350, 250);
            this.Name = "AddMacroForm";
            this.Text = "매크로 추가";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 컨트롤 추가
            Label lblKeyword = new Label();
            lblKeyword.Text = "음성 명령어:";
            lblKeyword.Location = new Point(20, 20);
            lblKeyword.Size = new Size(100, 20);
            this.Controls.Add(lblKeyword);

            txtKeyword = new TextBox();
            txtKeyword.Location = new Point(130, 20);
            txtKeyword.Size = new Size(200, 20);
            this.Controls.Add(txtKeyword);

            Label lblAction = new Label();
            lblAction.Text = "키 동작:";
            lblAction.Location = new Point(20, 60);
            lblAction.Size = new Size(100, 20);
            this.Controls.Add(lblAction);

            txtAction = new TextBox();
            txtAction.Location = new Point(130, 60);
            txtAction.Size = new Size(200, 20);
            this.Controls.Add(txtAction);

            Label lblActionInfo = new Label();
            lblActionInfo.Text = "예: CTRL+C, ALT+TAB, F5 등";
            lblActionInfo.Location = new Point(130, 85);
            lblActionInfo.Size = new Size(200, 20);
            lblActionInfo.ForeColor = Color.Gray;
            lblActionInfo.Font = new Font(lblActionInfo.Font.FontFamily, lblActionInfo.Font.Size - 1);
            this.Controls.Add(lblActionInfo);
            
            // 녹음 상태를 표시할 프로그레스 바 추가
            progressRecording = new ProgressBar();
            progressRecording.Location = new Point(130, 115);
            progressRecording.Size = new Size(200, 15);
            progressRecording.Minimum = 0;
            progressRecording.Maximum = 100;
            progressRecording.Value = 0;
            this.Controls.Add(progressRecording);
            
            // 녹음 상태 표시 레이블
            lblRecordingStatus = new Label();
            lblRecordingStatus.Text = "준비";
            lblRecordingStatus.Location = new Point(130, 135);
            lblRecordingStatus.Size = new Size(200, 20);
            lblRecordingStatus.ForeColor = Color.Gray;
            this.Controls.Add(lblRecordingStatus);

            btnRecord = new Button();
            btnRecord.Text = "음성 녹음";
            btnRecord.Location = new Point(20, 120);
            btnRecord.Size = new Size(90, 30);
            btnRecord.Click += BtnRecord_Click;
            this.Controls.Add(btnRecord);

            Button btnOk = new Button();
            btnOk.Text = "확인";
            btnOk.Location = new Point(160, 170);
            btnOk.Size = new Size(80, 30);
            btnOk.Click += (sender, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtKeyword.Text))
                {
                    MessageBox.Show("음성 명령어를 입력하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (string.IsNullOrWhiteSpace(txtAction.Text))
                {
                    MessageBox.Show("키 동작을 입력하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Keyword = txtKeyword.Text.Trim();
                KeyAction = txtAction.Text.Trim();
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            this.Controls.Add(btnOk);

            Button btnCancel = new Button();
            btnCancel.Text = "취소";
            btnCancel.Location = new Point(250, 170);
            btnCancel.Size = new Size(80, 30);
            btnCancel.Click += (sender, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };
            this.Controls.Add(btnCancel);

            this.ResumeLayout(false);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 폼이 닫힐 때 리소스 정리
            cancellationTokenSource?.Cancel();
            cancellationTokenSource?.Dispose();
            audioRecorder?.Dispose();
            
            base.OnFormClosing(e);
        }
    }
} 