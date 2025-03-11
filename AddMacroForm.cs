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
using System.IO;
using NAudio.Wave;
using VoiceMacro.Services;

namespace VoiceMacro
{
    /// <summary>
    /// 음성 매크로 추가를 위한 폼 클래스입니다.
    /// 음성 녹음, 인식, 매크로 설정 기능을 제공합니다.
    /// </summary>
    public partial class AddMacroForm : Form
    {
        /// <summary>
        /// 매크로의 음성 키워드입니다. 이 키워드가 인식되면 매크로가 실행됩니다.
        /// </summary>
        public string Keyword { get; private set; }

        /// <summary>
        /// 매크로가 실행될 때 수행할 키보드 동작입니다.
        /// </summary>
        public string KeyAction { get; private set; }

        /// <summary>
        /// 애플리케이션 설정 정보를 담고 있는 객체입니다.
        /// </summary>
        private readonly AppSettings settings;

        /// <summary>
        /// 오디오 녹음 기능을 제공하는 서비스입니다.
        /// </summary>
        private readonly AudioRecordingService audioRecorder;

        /// <summary>
        /// OpenAI API를 사용한 음성 인식 서비스입니다.
        /// </summary>
        private readonly OpenAIService openAIService;

        /// <summary>
        /// 음성 인식 서비스입니다.
        /// </summary>
        private readonly VoiceRecognitionService voiceRecognitionService;

        /// <summary>
        /// 음성 키워드를 입력하는 텍스트박스입니다.
        /// </summary>
        private TextBox txtKeyword;

        /// <summary>
        /// 키보드 동작을 입력하는 텍스트박스입니다.
        /// </summary>
        private TextBox txtAction;

        /// <summary>
        /// 음성 녹음 시작/중지 버튼입니다.
        /// </summary>
        private Button btnRecord;

        /// <summary>
        /// 녹음된 오디오를 파일로 저장하는 버튼입니다.
        /// </summary>
        private Button btnSaveAudio;

        /// <summary>
        /// 녹음 작업 취소를 위한 토큰 소스입니다.
        /// </summary>
        private CancellationTokenSource cancellationTokenSource;

        /// <summary>
        /// 녹음 중 오디오 레벨을 시각적으로 표시하는 프로그레스 바입니다.
        /// </summary>
        private ProgressBar progressRecording;

        /// <summary>
        /// 녹음 상태를 표시하는 레이블입니다.
        /// </summary>
        private Label lblRecordingStatus;

        /// <summary>
        /// 현재 녹음 중인지 여부를 나타내는 플래그입니다.
        /// </summary>
        private bool isRecording = false;

        /// <summary>
        /// 마지막으로 녹음된 오디오 데이터를 저장합니다.
        /// </summary>
        private byte[] lastRecordedAudio = null;

        /// <summary>
        /// 액션 타입 선택을 위한 콤보박스 추가
        /// </summary>
        private ComboBox cmbActionType;
        /// <summary>
        /// 액션 파라미터 입력을 위한 텍스트 박스 추가
        /// </summary>
        private TextBox txtActionParam;
        /// <summary>
        /// 선택된 액션 타입 (기본값: Default)
        /// </summary>
        private MacroActionType selectedActionType = MacroActionType.Default;
        /// <summary>
        /// 선택된 액션 파라미터 (기본값: 0)
        /// </summary>
        private int selectedActionParam = 0;
        
        /// <summary>
        /// 편집 모드인지 여부
        /// </summary>
        private bool isEditMode = false;
        
        /// <summary>
        /// 원본 매크로 키워드 (편집 모드일 때 사용)
        /// </summary>
        private string originalKeyword = string.Empty;

        /// <summary>
        /// 매크로 추가 폼의 생성자입니다.
        /// </summary>
        /// <param name="voiceRecognitionService">음성 인식 서비스 인스턴스</param>
        public AddMacroForm(VoiceRecognitionService voiceRecognitionService)
        {
            this.voiceRecognitionService = voiceRecognitionService;
            this.audioRecorder = new AudioRecordingService();
            this.settings = new AppSettings(); // 기본 설정 사용 (설정 파일에서 로드)
            
            // OpenAI API 서비스 초기화
            if (!string.IsNullOrEmpty(settings.OpenAIApiKey))
            {
                openAIService = new OpenAIService(settings.OpenAIApiKey);
            }
            
            // 오디오 레코더 이벤트 구독
            audioRecorder.RecordingStatusChanged += AudioRecorder_RecordingStatusChanged;
            audioRecorder.AudioLevelChanged += AudioRecorder_AudioLevelChanged;
            
            InitializeComponent();
            
            // 음성 인식 팁 버튼 추가
            AddVoiceRecognitionTipButton();
        }
        
        /// <summary>
        /// 매크로 편집 모드를 위한 생성자입니다.
        /// </summary>
        /// <param name="voiceRecognitionService">음성 인식 서비스 인스턴스</param>
        /// <param name="keyword">편집할 매크로 키워드</param>
        /// <param name="keyAction">편집할 매크로 키 동작</param>
        /// <param name="actionType">편집할 매크로 액션 타입</param>
        /// <param name="actionParam">편집할 매크로 액션 파라미터</param>
        public AddMacroForm(VoiceRecognitionService voiceRecognitionService, string keyword, string keyAction, 
                            MacroActionType actionType, int actionParam)
            : this(voiceRecognitionService)
        {
            // 편집 모드 설정
            isEditMode = true;
            originalKeyword = keyword;
            selectedActionType = actionType;
            SelectedActionType = actionType;
            selectedActionParam = actionParam;
            SelectedActionParam = actionParam;
            
            // 폼 제목 변경
            this.Text = "매크로 편집";
            
            // 폼이 로드된 후 값 설정 (컨트롤 초기화 후)
            this.Load += (sender, e) =>
            {
                txtKeyword.Text = keyword;
                txtAction.Text = keyAction;
                cmbActionType.SelectedIndex = (int)actionType;
                
                // 액션 타입에 따라 파라미터 설정
                txtActionParam.Text = actionParam.ToString();
                
                // 확인 버튼 텍스트 변경
                Button okButton = this.Controls.Find("btnOk", true).FirstOrDefault() as Button;
                if (okButton != null)
                {
                    okButton.Text = "저장";
                }
            };
        }

        /// <summary>
        /// 오디오 녹음 상태가 변경될 때 호출되는 이벤트 핸들러입니다.
        /// </summary>
        /// <param name="sender">이벤트 발생 객체</param>
        /// <param name="status">변경된 상태 메시지</param>
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

        /// <summary>
        /// 오디오 레벨이 변경될 때 호출되는 이벤트 핸들러입니다.
        /// </summary>
        /// <param name="sender">이벤트 발생 객체</param>
        /// <param name="level">현재 오디오 레벨 (dB)</param>
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

        /// <summary>
        /// 음성 녹음 버튼 클릭 이벤트 핸들러입니다.
        /// 버튼 상태에 따라 녹음을 시작하거나 중지합니다.
        /// </summary>
        /// <param name="sender">이벤트 발생 객체</param>
        /// <param name="e">이벤트 인자</param>
        private async void BtnRecord_Click(object sender, EventArgs e)
        {
            // 이미 녹음 중이면 녹음 중지
            if (isRecording)
            {
                // 녹음 중지
                try
                {
                    cancellationTokenSource?.Cancel();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"녹음 취소 중 오류: {ex.Message}");
                }
                finally
                {
                    btnRecord.Text = "음성 녹음";
                    isRecording = false;
                    lblRecordingStatus.Text = "취소됨";
                }
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
                cancellationTokenSource = new CancellationTokenSource();
                btnRecord.Text = "녹음 중지";
                isRecording = true;
                lblRecordingStatus.Text = "말씀하세요...";
                btnSaveAudio.Enabled = false;

                // 레코딩 시작 시 이전 인식 결과 초기화
                ClearRecognitionLabels();
                
                // 음성 인식 예시 표시 - 사용자에게 어떤 종류의 명령어가 잘 인식되는지 보여줍니다
                ShowRecognitionExamples();

                try
                {
                    // 녹음 시작 및 결과 처리
                    lastRecordedAudio = await audioRecorder.RecordSpeechAsync(cancellationTokenSource.Token);
                    
                    if (lastRecordedAudio != null && lastRecordedAudio.Length > 0)
                    {
                        btnSaveAudio.Enabled = true;
                        
                        try
                        {
                            // 음성 처리 중임을 표시
                            lblRecordingStatus.Text = "음성 분석 중...";
                            Application.DoEvents();
                            
                            string recognizedText = "";
                            
                            // OpenAI API 사용 (설정에 따라)
                            if (settings != null && settings.UseOpenAI && openAIService != null)
                            {
                                recognizedText = await openAIService.TranscribeAudioAsync(
                                    lastRecordedAudio, settings?.WhisperLanguage ?? "ko");
                            }
                            else
                            {
                                // VoiceRecognitionService의 로컬 Whisper 프로세서를 활용
                                recognizedText = await voiceRecognitionService.RecognizeAudioAsync(
                                    lastRecordedAudio, settings.WhisperLanguage);
                            }
                            
                            if (!string.IsNullOrWhiteSpace(recognizedText))
                            {
                                // 인식된 텍스트에서 키워드 추출
                                string keyword = ExtractKeyword(recognizedText);
                                
                                // 키워드 필드에는 입력하지 않고 별도 표시만 함
                                // txtKeyword.Text = keyword;
                                
                                // 상태 레이블에 인식된 원본 텍스트 표시 (최대 50자)
                                string displayText = recognizedText.Length > 50 
                                    ? recognizedText.Substring(0, 47) + "..." 
                                    : recognizedText;
                                
                                lblRecordingStatus.Text = "인식 완료";
                                
                                // 인식 결과 표시 (별도 팝업 또는 확장된 레이블)
                                ShowRecognitionResult(displayText);
                                
                                // 추출된 키워드를 표시하는 별도 레이블 추가
                                ShowExtractedKeyword(keyword);
                                
                                // 변환 과정 설명 표시
                                ShowTransformationProcess(recognizedText, keyword);
                            }
                            else
                            {
                                // 인식 실패 시 매크로 이름은 변경하지 않고 상태만 표시
                                lblRecordingStatus.Text = "인식 실패 (키워드를 직접 입력해주세요)";
                                txtKeyword.Focus(); // 사용자가 직접 입력할 수 있도록 포커스 이동
                                
                                // 이전 인식 결과 및 키워드 레이블 초기화
                                ClearRecognitionLabels();
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"음성 인식 중 오류: {ex.Message}");
                            lblRecordingStatus.Text = "인식 처리 오류";
                        }
                    }
                    else
                    {
                        // 오디오 데이터가 없는 경우 (마이크 감지 실패 등)
                        lblRecordingStatus.Text = "녹음 데이터 없음";
                        btnSaveAudio.Enabled = false;
                        lastRecordedAudio = null;
                    }
                }
                catch (OperationCanceledException)
                {
                    // 취소된 경우 - 정상적인 처리
                    lblRecordingStatus.Text = "녹음 취소됨";
                    btnSaveAudio.Enabled = false;
                    lastRecordedAudio = null;
                    
                    // 인식 결과 및 키워드 레이블 초기화
                    ClearRecognitionLabels();
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
                btnSaveAudio.Enabled = false;
                lastRecordedAudio = null;
            }
            finally
            {
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
            }
        }

        /// <summary>
        /// 폼 컴포넌트들을 초기화합니다.
        /// </summary>
        private void InitializeComponent()
        {
            // 폼 설정
            this.ClientSize = new System.Drawing.Size(750, 500); // 더 넓게 조정
            this.Name = "AddMacroForm";
            this.Text = "매크로 추가";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // 탭 컨트롤 추가
            TabControl tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;
            tabControl.Padding = new Point(10, 10);
            
            // 기본 탭 페이지
            TabPage basicTabPage = new TabPage("기본 설정");
            basicTabPage.Padding = new Padding(10);
            
            // 고급 탭 페이지
            TabPage advancedTabPage = new TabPage("고급 설정");
            advancedTabPage.Padding = new Padding(10);
            
            // 음성 인식 탭 페이지
            TabPage voiceTabPage = new TabPage("음성 인식");
            voiceTabPage.Padding = new Padding(10);
            
            // 기본 탭 페이지 컨트롤 추가
            // 1. 기본 정보 그룹 박스
            GroupBox basicInfoGroup = new GroupBox();
            basicInfoGroup.Text = "매크로 정보";
            basicInfoGroup.Size = new Size(680, 220); // 크기 증가
            basicInfoGroup.Location = new Point(10, 10);
            basicTabPage.Controls.Add(basicInfoGroup);
            
            // 키워드 레이블과 텍스트 박스
            Label lblKeyword = new Label();
            lblKeyword.Text = "음성 명령어:";
            lblKeyword.Location = new Point(20, 30);
            lblKeyword.Size = new Size(100, 20);
            basicInfoGroup.Controls.Add(lblKeyword);

            txtKeyword = new TextBox();
            txtKeyword.Location = new Point(130, 30);
            txtKeyword.Size = new Size(250, 20);
            basicInfoGroup.Controls.Add(txtKeyword);

            // 키 동작 레이블과 텍스트 박스
            Label lblAction = new Label();
            lblAction.Text = "키 동작:";
            lblAction.Location = new Point(20, 60);
            lblAction.Size = new Size(100, 20);
            basicInfoGroup.Controls.Add(lblAction);

            txtAction = new TextBox();
            txtAction.Location = new Point(130, 60);
            txtAction.Size = new Size(250, 20);
            basicInfoGroup.Controls.Add(txtAction);

            // 키 동작 안내 레이블
            Label lblActionInfo = new Label();
            lblActionInfo.Text = "예: CTRL+C, ALT+TAB, F5 등";
            lblActionInfo.Location = new Point(130, 85);
            lblActionInfo.Size = new Size(200, 20);
            lblActionInfo.ForeColor = Color.Gray;
            lblActionInfo.Font = new Font(lblActionInfo.Font.FontFamily, lblActionInfo.Font.Size - 1);
            basicInfoGroup.Controls.Add(lblActionInfo);
            
            // 가상 키패드 그룹박스
            GroupBox virtualKeypadGroup = new GroupBox();
            virtualKeypadGroup.Text = "가상 키패드";
            virtualKeypadGroup.Location = new Point(390, 20);
            virtualKeypadGroup.Size = new Size(280, 190);
            basicInfoGroup.Controls.Add(virtualKeypadGroup);
            
            // 키패드 패널 (FlowLayoutPanel 사용)
            FlowLayoutPanel keypadPanel = new FlowLayoutPanel();
            keypadPanel.Dock = DockStyle.Fill;
            keypadPanel.Padding = new Padding(5);
            keypadPanel.AutoScroll = true;
            virtualKeypadGroup.Controls.Add(keypadPanel);
            
            // 자주 사용되는 특수 키 버튼들 추가
            string[] specialKeys = new string[] {
                "{ENTER}", "{ESC}", "{TAB}", "{SPACE}", 
                "{BACKSPACE}", "{DELETE}", "{INSERT}",
                "{HOME}", "{END}", "{PGUP}", "{PGDN}",
                "{UP}", "{DOWN}", "{LEFT}", "{RIGHT}",
                "{F1}", "{F2}", "{F3}", "{F4}", "{F5}", 
                "{F6}", "{F7}", "{F8}", "{F9}", "{F10}",
                "{F11}", "{F12}", "{PRTSC}", "{BREAK}", "{PAUSE}",
                "{PAUP}", "{PADN}" // 요청한 특수 키 추가
            };
            
            foreach (string key in specialKeys)
            {
                Button keyButton = new Button();
                keyButton.Text = key.Replace("{", "").Replace("}", "");
                keyButton.Tag = key;
                keyButton.Size = new Size(70, 30);
                keyButton.Margin = new Padding(3);
                keyButton.BackColor = Color.LightBlue;
                keyButton.Click += KeyButton_Click;
                keypadPanel.Controls.Add(keyButton);
            }
            
            // 조합 키 섹션 추가
            Label lblModifierKeys = new Label();
            lblModifierKeys.Text = "조합 키";
            lblModifierKeys.AutoSize = true;
            lblModifierKeys.Margin = new Padding(3, 10, 3, 3);
            keypadPanel.Controls.Add(lblModifierKeys);
            
            // 새 패널로 조합 키 감싸기
            FlowLayoutPanel modifierPanel = new FlowLayoutPanel();
            modifierPanel.Size = new Size(280, 35);
            modifierPanel.Margin = new Padding(0);
            keypadPanel.Controls.Add(modifierPanel);
            
            // 조합 키 버튼 추가
            string[] modifierKeys = new string[] { "CTRL+", "ALT+", "SHIFT+", "WIN+" };
            foreach (string key in modifierKeys)
            {
                Button keyButton = new Button();
                keyButton.Text = key;
                keyButton.Tag = key;
                keyButton.Size = new Size(65, 30);
                keyButton.Margin = new Padding(3);
                keyButton.BackColor = Color.LightGreen;
                keyButton.Click += KeyButton_Click;
                modifierPanel.Controls.Add(keyButton);
            }
            
            // 액션 타입 그룹
            GroupBox actionTypeGroup = new GroupBox();
            actionTypeGroup.Text = "액션 타입";
            actionTypeGroup.Size = new Size(400, 120);
            actionTypeGroup.Location = new Point(10, 240);
            basicTabPage.Controls.Add(actionTypeGroup);
            
            // 액션 타입 선택 레이블과 콤보박스
            Label lblActionType = new Label();
            lblActionType.Text = "액션 타입:";
            lblActionType.Location = new Point(20, 30);
            lblActionType.Size = new Size(100, 20);
            actionTypeGroup.Controls.Add(lblActionType);
            
            cmbActionType = new ComboBox();
            cmbActionType.Location = new Point(130, 30);
            cmbActionType.Size = new Size(250, 20);
            cmbActionType.DropDownStyle = ComboBoxStyle.DropDownList;
            
            // 콤보박스에 액션 타입 추가
            cmbActionType.Items.Add(new { Text = "기본 - 한 번 입력", Value = MacroActionType.Default });
            cmbActionType.Items.Add(new { Text = "토글 - 키를 누르고 떼는 동작 전환", Value = MacroActionType.Toggle });
            cmbActionType.Items.Add(new { Text = "반복 - n번 반복", Value = MacroActionType.Repeat });
            cmbActionType.Items.Add(new { Text = "홀드 - n초 동안 누르기", Value = MacroActionType.Hold });
            cmbActionType.Items.Add(new { Text = "터보 - n밀리초 간격으로 연타", Value = MacroActionType.Turbo });
            cmbActionType.Items.Add(new { Text = "콤보 - 여러 키 순차 입력", Value = MacroActionType.Combo });
            
            // 표시 형식 지정
            cmbActionType.DisplayMember = "Text";
            cmbActionType.ValueMember = "Value";
            
            // 기본값 선택
            cmbActionType.SelectedIndex = 0;
            
            // 이벤트 연결
            cmbActionType.SelectedIndexChanged += CmbActionType_SelectedIndexChanged;
            
            actionTypeGroup.Controls.Add(cmbActionType);
            
            // 액션 파라미터 레이블과 텍스트박스
            Label lblActionParam = new Label();
            lblActionParam.Text = "파라미터:";
            lblActionParam.Location = new Point(20, 70);
            lblActionParam.Size = new Size(100, 20);
            actionTypeGroup.Controls.Add(lblActionParam);
            
            txtActionParam = new TextBox();
            txtActionParam.Location = new Point(130, 70);
            txtActionParam.Size = new Size(100, 20);
            txtActionParam.Text = "0";
            txtActionParam.Enabled = false; // 기본값은 비활성화
            txtActionParam.TextChanged += TxtActionParam_TextChanged;
            actionTypeGroup.Controls.Add(txtActionParam);
            
            // 액션 파라미터 도움말
            Label lblActionParamInfo = new Label();
            lblActionParamInfo.Text = "반복 횟수, 지속 시간(초) 등";
            lblActionParamInfo.Location = new Point(240, 70);
            lblActionParamInfo.Size = new Size(200, 20);
            lblActionParamInfo.ForeColor = Color.Gray;
            lblActionParamInfo.Font = new Font(lblActionParamInfo.Font.FontFamily, lblActionParamInfo.Font.Size - 1);
            actionTypeGroup.Controls.Add(lblActionParamInfo);
            
            // 음성 탭 구성 - 녹음 그룹
            GroupBox recordingGroup = new GroupBox();
            recordingGroup.Text = "음성 녹음";
            recordingGroup.Size = new Size(400, 280);
            recordingGroup.Location = new Point(10, 10);
            voiceTabPage.Controls.Add(recordingGroup);
            
            // 녹음 상태 레이블
            lblRecordingStatus = new Label();
            lblRecordingStatus.Text = "녹음 준비됨";
            lblRecordingStatus.Location = new Point(20, 30);
            lblRecordingStatus.Size = new Size(360, 20);
            lblRecordingStatus.TextAlign = ContentAlignment.MiddleCenter;
            lblRecordingStatus.Font = new Font(lblRecordingStatus.Font, FontStyle.Bold);
            recordingGroup.Controls.Add(lblRecordingStatus);
            
            // 녹음 버튼
            btnRecord = new Button();
            btnRecord.Text = "음성 녹음 시작";
            btnRecord.Location = new Point(20, 60);
            btnRecord.Size = new Size(360, 40);
            btnRecord.BackColor = Color.LightBlue;
            btnRecord.Click += BtnRecord_Click;
            recordingGroup.Controls.Add(btnRecord);
            
            // 오디오 저장 버튼
            btnSaveAudio = new Button();
            btnSaveAudio.Text = "녹음된 오디오 저장";
            btnSaveAudio.Location = new Point(20, 110);
            btnSaveAudio.Size = new Size(360, 30);
            btnSaveAudio.Enabled = false;
            btnSaveAudio.Click += BtnSaveAudio_Click;
            recordingGroup.Controls.Add(btnSaveAudio);
            
            // 녹음 진행 막대
            progressRecording = new ProgressBar();
            progressRecording.Location = new Point(20, 150);
            progressRecording.Size = new Size(360, 20);
            progressRecording.Style = ProgressBarStyle.Continuous;
            progressRecording.Minimum = 0;
            progressRecording.Maximum = 100;
            progressRecording.Value = 0;
            recordingGroup.Controls.Add(progressRecording);
            
            // 진행 표시 레이블
            Label lblProgress = new Label();
            lblProgress.Text = "음성 레벨:";
            lblProgress.Location = new Point(20, 175);
            lblProgress.Size = new Size(80, 20);
            recordingGroup.Controls.Add(lblProgress);
            
            // 탭 추가
            tabControl.TabPages.Add(basicTabPage);
            tabControl.TabPages.Add(advancedTabPage);
            tabControl.TabPages.Add(voiceTabPage);
            
            // 폼에 탭 컨트롤 추가
            this.Controls.Add(tabControl);
            
            // 버튼 패널
            Panel buttonPanel = new Panel();
            buttonPanel.Dock = DockStyle.Bottom;
            buttonPanel.Height = 50;
            this.Controls.Add(buttonPanel);
            
            // OK 버튼
            Button btnOk = new Button();
            btnOk.Text = "확인";
            btnOk.Location = new Point(550, 10);
            btnOk.Size = new Size(80, 30);
            btnOk.Click += (sender, e) =>
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            buttonPanel.Controls.Add(btnOk);
            
            // 취소 버튼
            Button btnCancel = new Button();
            btnCancel.Text = "취소";
            btnCancel.Location = new Point(640, 10);
            btnCancel.Size = new Size(80, 30);
            btnCancel.Click += (sender, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };
            buttonPanel.Controls.Add(btnCancel);
        }

        /// <summary>
        /// 폼이 닫힐 때 호출되는 이벤트 핸들러입니다.
        /// 사용 중인 리소스를 정리합니다.
        /// </summary>
        /// <param name="e">폼 닫기 이벤트 인자</param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            
            if (DialogResult == DialogResult.OK)
            {
                Keyword = txtKeyword.Text.Trim();
                KeyAction = txtAction.Text.Trim();
                
                // DialogResult가 OK일 때만 MacroService로 전달할 액션 타입과 파라미터 저장
                SelectedActionType = selectedActionType;
                SelectedActionParam = selectedActionParam;
            }
            
            // 폼이 닫힐 때 리소스 정리
            try
            {
                // 녹음 중이면 녹음 중지
                if (isRecording)
                {
                    isRecording = false;
                    cancellationTokenSource?.Cancel();
                }
                
                // 리소스 정리
                cancellationTokenSource?.Cancel();
                cancellationTokenSource?.Dispose();
                audioRecorder?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"폼 닫기 중 리소스 정리 오류: {ex.Message}");
                // 예외가 발생해도 계속 진행
            }
        }

        /// <summary>
        /// 선택된 액션 타입을 가져옵니다.
        /// </summary>
        public MacroActionType SelectedActionType { get; private set; } = MacroActionType.Default;
        
        /// <summary>
        /// 선택된 액션 파라미터를 가져옵니다.
        /// </summary>
        public int SelectedActionParam { get; private set; } = 0;
        
        /// <summary>
        /// 원본 매크로 키워드를 가져옵니다. (편집 모드에서만 의미 있음)
        /// </summary>
        public string OriginalKeyword => originalKeyword;
        
        /// <summary>
        /// 편집 모드인지 여부를 가져옵니다.
        /// </summary>
        public bool IsEditMode => isEditMode;

        /// <summary>
        /// 오디오 저장 버튼 클릭 이벤트 핸들러입니다.
        /// 녹음된 오디오를 WAV 파일로 저장합니다.
        /// </summary>
        /// <param name="sender">이벤트 발생 객체</param>
        /// <param name="e">이벤트 인자</param>
        private async void BtnSaveAudio_Click(object sender, EventArgs e)
        {
            // 저장할 오디오 데이터 확인
            if (lastRecordedAudio == null || lastRecordedAudio.Length == 0)
            {
                MessageBox.Show("저장할 녹음 데이터가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 파일 저장 대화상자 표시
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "WAV 파일 (*.wav)|*.wav|모든 파일 (*.*)|*.*";
            saveFileDialog.DefaultExt = "wav";
            saveFileDialog.Title = "녹음 파일 저장";
            saveFileDialog.FileName = $"녹음_{DateTime.Now:yyyyMMdd_HHmmss}.wav";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 오디오 파일 저장
                    await SaveWavFileAsync(lastRecordedAudio, saveFileDialog.FileName);
                    lblRecordingStatus.Text = "파일 저장 완료";
                    MessageBox.Show($"파일이 저장되었습니다: {saveFileDialog.FileName}", "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"파일 저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// 오디오 데이터를 WAV 파일로 저장합니다.
        /// </summary>
        /// <param name="audioData">저장할 오디오 데이터</param>
        /// <param name="filePath">저장할 파일 경로</param>
        /// <returns>작업 완료를 나타내는 Task</returns>
        private async Task SaveWavFileAsync(byte[] audioData, string filePath)
        {
            await Task.Run(() =>
            {
                try
                {
                    // 원본 데이터가 이미 WAV 포맷인지 확인
                    if (IsWavFormat(audioData))
                    {
                        // 이미 WAV 형식이면 바로 저장
                        File.WriteAllBytes(filePath, audioData);
                    }
                    else
                    {
                        // 오디오 데이터를 WAV 포맷으로 변환하여 저장
                        using (MemoryStream sourceStream = new MemoryStream(audioData))
                        using (WaveFileWriter waveWriter = new WaveFileWriter(filePath, new WaveFormat(16000, 16, 1)))
                        {
                            // 오디오 데이터를 WAV 형식으로 변환하여 저장
                            byte[] buffer = new byte[4096];
                            int bytesRead;
                            
                            sourceStream.Position = 0;
                            while ((bytesRead = sourceStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                waveWriter.Write(buffer, 0, bytesRead);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WAV 파일 저장 오류: {ex.Message}");
                    throw; // 상위 catch 블록에서 처리하도록 예외 다시 throw
                }
            });
        }

        /// <summary>
        /// 바이트 배열이 WAV 형식인지 확인합니다.
        /// </summary>
        /// <param name="data">확인할 바이트 배열</param>
        /// <returns>WAV 형식이면 true, 아니면 false</returns>
        private bool IsWavFormat(byte[] data)
        {
            // WAV 파일 헤더 확인 (RIFF + WAVE 시그니처)
            if (data.Length < 12) return false;
            
            try
            {
                string signature = System.Text.Encoding.ASCII.GetString(data, 0, 4);
                string format = System.Text.Encoding.ASCII.GetString(data, 8, 4);
                
                return signature == "RIFF" && format == "WAVE";
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 인식된 음성 텍스트에서 핵심 키워드를 추출합니다.
        /// </summary>
        /// <param name="text">인식된 전체 텍스트</param>
        /// <returns>추출된 키워드</returns>
        private string ExtractKeyword(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // 문장 정리 (원본 보존)
            string originalText = text;
            text = text.Trim();
            
            // 문장 부호 제거
            text = text.Replace(".", "").Replace(",", "").Replace("!", "").Replace("?", "");
            
            // 불필요한 대화체 표현 제거
            string[] unnecessaryPhrases = new string[] 
            { 
                // 일반 간투사
                "음", "그", "저", "어", "아", "이제", "이거", "저거", "그거",
                
                // 요청 표현
                "해줘", "해 줘", "해주세요", "해 주세요", "부탁해", "부탁합니다", 
                "실행", "실행해", "실행해줘", "실행 해줘", "실행해 줘", "실행 해 줘",
                
                // 매크로 관련 표현
                "매크로", "매크로로", "기능", "기능을", "단축키", "단축키로",
                "명령", "명령어", "로 설정", "설정해줘", "설정 해줘",
                
                // 추가 대화체
                "좀", "잠깐", "지금", "빨리", "제발", "가능하면"
            };
            
            // 정규 표현식에서 에러를 일으킬 수 있는 특수 문자 이스케이프
            for (int i = 0; i < unnecessaryPhrases.Length; i++)
            {
                unnecessaryPhrases[i] = System.Text.RegularExpressions.Regex.Escape(unnecessaryPhrases[i]);
            }
            
            // 불필요한 표현 제거
            foreach (var phrase in unnecessaryPhrases)
            {
                // 단어 전체가 일치할 때만 공백으로 대체 (부분 일치 방지)
                text = System.Text.RegularExpressions.Regex.Replace(
                    text, 
                    $@"\b{phrase}\b", 
                    "", 
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                );
            }
            
            // 한국어 조사 제거 (을/를, 이/가, 은/는 등)
            string[] koreanParticles = new string[] 
            {
                "을", "를", "이", "가", "은", "는", "에", "에서", "로", "으로", 
                "와", "과", "랑", "이랑", "하고", "에게", "한테", "께", "의"
            };
            
            // 조사 제거
            foreach (var particle in koreanParticles)
            {
                // 조사는 단어 뒤에 붙으므로 공백 또는 문장 끝 앞에 있을 때만 제거
                text = System.Text.RegularExpressions.Regex.Replace(
                    text, 
                    $@"{particle}(\s|$)", 
                    " ", 
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                );
            }
            
            // 앞뒤 공백과 중복 공백 제거
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ").Trim();
            
            // 일정 길이 이상이면 잘라내기
            const int maxKeywordLength = 15;
            if (text.Length > maxKeywordLength)
            {
                // 공백으로 분리하여 첫 몇 단어만 사용
                string[] words = text.Split(' ');
                text = string.Join(" ", words.Take(3)); // 최대 3개 단어만 사용
                
                // 여전히 길다면 그냥 잘라내기
                if (text.Length > maxKeywordLength)
                {
                    text = text.Substring(0, maxKeywordLength);
                }
            }
            
            // 빈 문자열이 된 경우 원본 텍스트의 일부 반환
            if (string.IsNullOrWhiteSpace(text))
            {
                // 원본 텍스트에서 첫 단어 추출
                string[] words = originalText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length > 0)
                {
                    // 첫 단어가 너무 길면 자르기
                    string firstWord = words[0];
                    if (firstWord.Length > maxKeywordLength)
                    {
                        firstWord = firstWord.Substring(0, maxKeywordLength);
                    }
                    return firstWord;
                }
                return originalText.Length > maxKeywordLength ? originalText.Substring(0, maxKeywordLength) : originalText;
            }
            
            return text;
        }

        /// <summary>
        /// 인식 결과를 사용자에게 표시합니다.
        /// </summary>
        /// <param name="recognizedText">인식된 텍스트</param>
        private void ShowRecognitionResult(string recognizedText)
        {
            // 탭 컨트롤의 "음성 인식" 탭에서 "인식 결과" 레이블 찾기
            foreach (Control control in this.Controls)
            {
                if (control is TabControl tabControl)
                {
                    foreach (TabPage page in tabControl.TabPages)
                    {
                        if (page.Text == "음성 인식")
                        {
                            foreach (Control pageControl in page.Controls)
                            {
                                if (pageControl is GroupBox groupBox && groupBox.Text == "음성 인식 결과")
                                {
                                    foreach (Control groupControl in groupBox.Controls)
                                    {
                                        if (groupControl is Label label && label.Name == "lblRecognitionResult")
                                        {
                                            label.Text = $"인식 결과: {recognizedText}";
                                            
                                            // "음성 인식" 탭으로 이동하여 결과 표시
                                            tabControl.SelectedIndex = tabControl.TabPages.IndexOf(page);
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 추출된 키워드를 사용자에게 표시합니다.
        /// </summary>
        /// <param name="keyword">추출된 키워드</param>
        private void ShowExtractedKeyword(string keyword)
        {
            // 탭 컨트롤의 "음성 인식" 탭에서 "추출된 키워드" 레이블 찾기
            foreach (Control control in this.Controls)
            {
                if (control is TabControl tabControl)
                {
                    foreach (TabPage page in tabControl.TabPages)
                    {
                        if (page.Text == "음성 인식")
                        {
                            foreach (Control pageControl in page.Controls)
                            {
                                if (pageControl is GroupBox groupBox && groupBox.Text == "음성 인식 결과")
                                {
                                    foreach (Control groupControl in groupBox.Controls)
                                    {
                                        if (groupControl is Label label && label.Name == "lblExtractedKeyword")
                                        {
                                            label.Text = $"추출된 키워드: {keyword}";
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 인식 결과 및 키워드 레이블을 초기화합니다.
        /// </summary>
        private void ClearRecognitionLabels()
        {
            // 탭 컨트롤의 "음성 인식" 탭에서 결과 레이블들 초기화
            foreach (Control control in this.Controls)
            {
                if (control is TabControl tabControl)
                {
                    foreach (TabPage page in tabControl.TabPages)
                    {
                        if (page.Text == "음성 인식")
                        {
                            foreach (Control pageControl in page.Controls)
                            {
                                if (pageControl is GroupBox groupBox && groupBox.Text == "음성 인식 결과")
                                {
                                    foreach (Control groupControl in groupBox.Controls)
                                    {
                                        if (groupControl is Label label)
                                        {
                                            if (label.Name == "lblRecognitionResult" || 
                                                label.Name == "lblExtractedKeyword" || 
                                                label.Name == "lblTransformation")
                                            {
                                                label.Text = "";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// 음성에서 텍스트로의 변환 과정을 시각적으로 표시합니다.
        /// </summary>
        /// <param name="originalText">원본 인식 텍스트</param>
        /// <param name="extractedKeyword">추출된 키워드</param>
        private void ShowTransformationProcess(string originalText, string extractedKeyword)
        {
            // 변환 과정 설명 (예: 단어 추출 등)
            string explanation;
            if (originalText.Equals(extractedKeyword, StringComparison.OrdinalIgnoreCase))
            {
                explanation = "음성이 그대로 키워드로 사용됩니다.";
            }
            else if (originalText.Contains(" ") && originalText.Split(' ')[0].Equals(extractedKeyword, StringComparison.OrdinalIgnoreCase))
            {
                explanation = "첫 단어가 키워드로 추출되었습니다.";
            }
            else
            {
                explanation = "음성에서 핵심 키워드가 추출되었습니다.";
            }
            
            // 탭 컨트롤의 "음성 인식" 탭에서 "변환 과정" 레이블 찾기
            foreach (Control control in this.Controls)
            {
                if (control is TabControl tabControl)
                {
                    foreach (TabPage page in tabControl.TabPages)
                    {
                        if (page.Text == "음성 인식")
                        {
                            foreach (Control pageControl in page.Controls)
                            {
                                if (pageControl is GroupBox groupBox && groupBox.Text == "음성 인식 결과")
                                {
                                    foreach (Control groupControl in groupBox.Controls)
                                    {
                                        if (groupControl is Label label && label.Name == "lblTransformation")
                                        {
                                            label.Text = "« 변환 과정 » " + explanation;
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// 음성 인식 예시를 표시합니다.
        /// </summary>
        private void ShowRecognitionExamples()
        {
            // 랜덤하게 예시 선택
            string[] examples = new string[]
            {
                "« 음성 인식 예시 »\n" +
                "- 명확한 단어 사용: \"엑셀 시작\", \"메모장 열기\"\n" +
                "- 짧은 문장 사용: \"화면 캡처해\", \"음량 올려줘\"\n" +
                "- 천천히 또박또박: 인식률이 높아집니다",
                
                "« 효과적인 음성 명령어 »\n" +
                "- 단순 명령: \"창 닫기\", \"파일 저장\"\n" +
                "- 앱 실행: \"브라우저 실행\", \"계산기 열기\"\n" +
                "- 추가 팁: 주변 소음이 적을수록 정확합니다",
                
                "« 음성 인식 활용 팁 »\n" +
                "- 자주 쓰는 단축키를 음성으로 설정해보세요\n" +
                "- 특수문자보다 일반 단어가 인식이 잘 됩니다\n" +
                "- 마이크와 입 사이 거리는 10-20cm 정도가 적합합니다"
            };
            
            // 랜덤 인덱스 생성
            Random random = new Random();
            int index = random.Next(examples.Length);
            
            // 탭 컨트롤의 "음성 인식" 탭에서 예시 레이블 찾기
            foreach (Control control in this.Controls)
            {
                if (control is TabControl tabControl)
                {
                    foreach (TabPage page in tabControl.TabPages)
                    {
                        if (page.Text == "음성 인식")
                        {
                            foreach (Control pageControl in page.Controls)
                            {
                                if (pageControl is GroupBox groupBox && groupBox.Text == "음성 인식 팁")
                                {
                                    foreach (Control groupControl in groupBox.Controls)
                                    {
                                        if (groupControl is Label label && label.Name == "lblExamples")
                                        {
                                            // 선택된 예시 표시
                                            label.Text = examples[index];
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 음성 인식 팁 버튼을 추가합니다.
        /// </summary>
        private void AddVoiceRecognitionTipButton()
        {
            Button btnTip = new Button();
            btnTip.Text = "인식 팁";
            btnTip.Location = new Point(230, 310);
            btnTip.Size = new Size(80, 30);
            btnTip.ForeColor = Color.Purple;
            btnTip.Click += (sender, e) => 
            {
                MessageBox.Show(
                    "음성 인식 효율을 높이기 위한 팁:\n\n" +
                    "1. 조용한 환경에서 녹음하세요.\n" +
                    "2. 마이크에 가까이 대고 또박또박 말하세요.\n" +
                    "3. 긴 문장보다 짧고 명확한 키워드를 사용하세요.\n" +
                    "4. 비슷한 발음의 단어는 피하세요.\n" +
                    "5. 음성 인식 실패 시 다른 단어로 시도해보세요.\n" +
                    "6. 키워드로 추출된 단어가 적합하지 않다면 '키워드 복사' 버튼을 누른 후 직접 수정하세요.",
                    "음성 인식 도움말",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                
                // 예시 새로고침
                ShowRecognitionExamples();
            };
            
            this.Controls.Add(btnTip);
        }

        /// <summary>
        /// 액션 타입 선택 변경 시 이벤트 핸들러
        /// </summary>
        private void CmbActionType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 선택된 인덱스에 따라 액션 타입 설정
            selectedActionType = (MacroActionType)cmbActionType.SelectedIndex;
            SelectedActionType = selectedActionType;
            
            // 액션 타입에 따라 파라미터 텍스트박스 활성화 및 기본값 설정
            switch (selectedActionType)
            {
                case MacroActionType.Default:
                    txtActionParam.Enabled = false;
                    txtActionParam.Text = "0";
                    UpdateActionParamDescription("파라미터가 필요하지 않습니다");
                    break;
                
                case MacroActionType.Toggle:
                    txtActionParam.Enabled = false;
                    txtActionParam.Text = "0";
                    UpdateActionParamDescription("파라미터가 필요하지 않습니다");
                    break;
                
                case MacroActionType.Repeat:
                    txtActionParam.Enabled = true;
                    txtActionParam.Text = "3";
                    UpdateActionParamDescription("반복 횟수 (기본값: 3회)");
                    break;
                
                case MacroActionType.Hold:
                    txtActionParam.Enabled = true;
                    txtActionParam.Text = "1000";
                    UpdateActionParamDescription("키 유지 시간 (밀리초, 기본값: 1000ms = 1초)");
                    break;
                
                case MacroActionType.Turbo:
                    txtActionParam.Enabled = true;
                    txtActionParam.Text = "50";
                    UpdateActionParamDescription("연타 간격 (밀리초, 기본값: 50ms)");
                    break;
                
                case MacroActionType.Combo:
                    txtActionParam.Enabled = true;
                    txtActionParam.Text = "100";
                    UpdateActionParamDescription("키 입력 간격 (밀리초, 기본값: 100ms)");
                    break;
            }
            
            // 키 동작 안내 레이블 업데이트
            if (selectedActionType == MacroActionType.Combo)
            {
                UpdateKeyActionGuide("쉼표로 구분된 키 목록 입력 (예: CTRL+C, F5, ENTER)");
            }
            else
            {
                UpdateKeyActionGuide("예: CTRL+C, ALT+TAB, F5 등");
            }
        }
        
        /// <summary>
        /// 액션 파라미터 텍스트 변경 시 이벤트 핸들러
        /// </summary>
        private void TxtActionParam_TextChanged(object sender, EventArgs e)
        {
            // 파라미터 값 유효성 검사 및 설정
            if (int.TryParse(txtActionParam.Text, out int value))
            {
                selectedActionParam = value;
                SelectedActionParam = value;
            }
            else
            {
                // 숫자가 아닌 경우 기본값으로 설정
                if (string.IsNullOrWhiteSpace(txtActionParam.Text))
                {
                    selectedActionParam = 0;
                    SelectedActionParam = 0;
                }
                else
                {
                    // 유효한 숫자가 아닌 값이 입력된 경우 경고
                    MessageBox.Show("숫자만 입력해주세요.", "입력 오류", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    
                    // 액션 타입에 따른 기본값으로 복원
                    switch (selectedActionType)
                    {
                        case MacroActionType.Repeat:
                            txtActionParam.Text = "3";
                            break;
                        case MacroActionType.Hold:
                            txtActionParam.Text = "1000";
                            break;
                        case MacroActionType.Turbo:
                            txtActionParam.Text = "50";
                            break;
                        case MacroActionType.Combo:
                            txtActionParam.Text = "100";
                            break;
                        default:
                            txtActionParam.Text = "0";
                            break;
                    }
                }
            }
        }
        
        /// <summary>
        /// 액션 파라미터 설명을 업데이트합니다.
        /// </summary>
        /// <param name="description">새 설명 텍스트</param>
        private void UpdateActionParamDescription(string description)
        {
            foreach (Control control in this.Controls)
            {
                if (control is Label label && label.Location.Y == 165 && label.Location.X == 130)
                {
                    label.Text = description;
                    return;
                }
            }
        }
        
        /// <summary>
        /// 키 동작 안내 텍스트를 업데이트합니다.
        /// </summary>
        /// <param name="guide">새 안내 텍스트</param>
        private void UpdateKeyActionGuide(string guide)
        {
            foreach (Control control in this.Controls)
            {
                if (control is Label label && label.Location.Y == 85 && label.Location.X == 130)
                {
                    label.Text = guide;
                    return;
                }
            }
        }

        /// <summary>
        /// 키 버튼 클릭 이벤트 핸들러
        /// </summary>
        /// <param name="sender">이벤트 발생 객체</param>
        /// <param name="e">이벤트 인자</param>
        private void KeyButton_Click(object sender, EventArgs e)
        {
            Button button = sender as Button;
            string keyCode = button.Tag.ToString();
            
            // 텍스트 상자에 키 코드 추가
            // 커서 위치에 삽입 또는 선택된 텍스트 대체
            if (txtAction.SelectionLength > 0)
            {
                txtAction.SelectedText = keyCode;
            }
            else
            {
                txtAction.Text = txtAction.Text.Insert(txtAction.SelectionStart, keyCode);
                txtAction.SelectionStart += keyCode.Length;
            }
            
            // 포커스 유지
            txtAction.Focus();
        }
    }
} 