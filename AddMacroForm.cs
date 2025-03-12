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
        public MacroActionType SelectedActionType { get; private set; } = MacroActionType.Default;
        public int SelectedActionParam { get; private set; } = 0;
        public string OriginalKeyword { get; private set; }

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
        
        // 액션 타입 관련 컨트롤
        private ComboBox cmbActionType;
        private Label lblActionType;
        private Label lblActionParam;
        private NumericUpDown numActionParam;
        private Label lblParamDescription;
        
        // 키 버튼 관련 컨트롤
        private Button btnCtrl;
        private Button btnAlt;
        private Button btnShift;
        private Button btnWin;
        private bool isCtrlPressed = false;
        private bool isAltPressed = false;
        private bool isShiftPressed = false;
        private bool isWinPressed = false;
        
        // 2x2 레이아웃을 위한 새 컨트롤들
        private TableLayoutPanel mainLayout;
        private GroupBox[] panelGroups;
        private RichTextBox txtHelpInfo;
        private ToolTip formToolTip;
        private FlowLayoutPanel keyboardPanel;

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
            
            // 기존 키 액션이 있으면 파싱하여 버튼 상태 업데이트
            if (!string.IsNullOrEmpty(KeyAction))
            {
                ParseKeyAction(KeyAction);
            }
        }

        // 이 생성자는 매크로 편집에 사용됩니다
        public AddMacroForm(VoiceRecognitionService voiceRecognitionService, 
                              string keyword, string keyAction, 
                              MacroActionType actionType, int actionParameters)
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
            
            // 기존 매크로 정보 저장
            this.OriginalKeyword = keyword;
            this.Keyword = keyword;
            this.KeyAction = keyAction;
            this.SelectedActionType = actionType;
            this.SelectedActionParam = actionParameters;
            
            InitializeComponent();
            
            // 기존 값으로 UI 초기화
            txtKeyword.Text = keyword;
            txtAction.Text = keyAction;
            
            // cmbActionType이 초기화되었는지 확인 후 설정
            if (cmbActionType != null && cmbActionType.Items.Count > 0)
            {
                cmbActionType.SelectedIndex = (int)actionType;
                
                // numActionParam이 초기화되었는지 확인 후 설정
                if (numActionParam != null)
                {
                    numActionParam.Value = actionParameters;
                }
            }
            
            // 기존 키 액션이 있으면 파싱하여 버튼 상태 업데이트
            if (!string.IsNullOrEmpty(keyAction))
            {
                ParseKeyAction(keyAction);
            }
            
            // 타이틀 변경
            this.Text = "매크로 수정";
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
            
            // 폼 설정 - 더 넓고 여유롭게
            this.ClientSize = new System.Drawing.Size(800, 600);
            this.Name = "AddMacroForm";
            this.Text = "매크로 추가";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // 툴팁 컴포넌트 초기화
            formToolTip = new ToolTip();
            formToolTip.AutoPopDelay = 5000;
            formToolTip.InitialDelay = 500;
            formToolTip.ReshowDelay = 500;
            formToolTip.ShowAlways = true;
            
            // 메인 2x2 레이아웃 설정
            mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.RowCount = 2;
            mainLayout.ColumnCount = 2;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            mainLayout.Padding = new Padding(10);
            mainLayout.Margin = new Padding(0);
            this.Controls.Add(mainLayout);
            
            // 4개의 패널(GroupBox) 생성
            panelGroups = new GroupBox[4];
            string[] panelTitles = {
                "음성 명령어 및 키 동작", "키 조합 및 모디파이어",
                "액션 타입 설정", "도움말 및 정보"
            };
            
            for (int i = 0; i < 4; i++)
            {
                panelGroups[i] = new GroupBox();
                panelGroups[i].Text = panelTitles[i];
                panelGroups[i].Dock = DockStyle.Fill;
                panelGroups[i].Margin = new Padding(5);
                panelGroups[i].Padding = new Padding(10);
                panelGroups[i].Font = new Font(this.Font.FontFamily, 10F, FontStyle.Bold);
                
                // 그리드 위치 계산 (0,0), (0,1), (1,0), (1,1)
                int row = i / 2;
                int col = i % 2;
                mainLayout.Controls.Add(panelGroups[i], col, row);
            }
            
            // 1. 왼쪽 상단 패널 - 음성 명령어 및 키 동작 설정
            InitializeVoiceAndKeyPanel(panelGroups[0]);
            
            // 2. 오른쪽 상단 패널 - 키 조합 및 모디파이어
            InitializeKeyModifierPanel(panelGroups[1]);
            
            // 3. 왼쪽 하단 패널 - 액션 타입 설정
            InitializeActionTypePanel(panelGroups[2]);
            
            // 4. 오른쪽 하단 패널 - 도움말 및 버튼
            InitializeHelpAndButtonPanel(panelGroups[3]);
            
            // 액션 타입 초기화 - 여기서 호출
            InitializeActionTypes();
            
            this.ResumeLayout(false);
        }
        
        /// <summary>
        /// 왼쪽 상단 패널 - 음성 명령어 및 키 동작 설정 초기화
        /// </summary>
        private void InitializeVoiceAndKeyPanel(GroupBox panel)
        {
            // 패널 내 컨트롤 배치를 위한 테이블 레이아웃
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.RowCount = 5;
            layout.ColumnCount = 2;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            panel.Controls.Add(layout);
            
            // 음성 명령어 라벨 및 텍스트박스
            Label lblKeyword = new Label();
            lblKeyword.Text = "음성 명령어:";
            lblKeyword.Dock = DockStyle.Fill;
            lblKeyword.TextAlign = ContentAlignment.MiddleLeft;
            lblKeyword.Font = new Font(this.Font.FontFamily, 9F);
            layout.Controls.Add(lblKeyword, 0, 0);
            
            txtKeyword = new TextBox();
            txtKeyword.Dock = DockStyle.Fill;
            txtKeyword.Margin = new Padding(3, 5, 3, 3);
            txtKeyword.Font = new Font(this.Font.FontFamily, 10F);
            layout.Controls.Add(txtKeyword, 1, 0);
            formToolTip.SetToolTip(txtKeyword, "음성으로 인식할 명령어를 입력하세요. 음성 녹음 버튼을 눌러 직접 녹음할 수도 있습니다.");
            
            // 키 동작 라벨 및 텍스트박스
            Label lblAction = new Label();
            lblAction.Text = "키 동작:";
            lblAction.Dock = DockStyle.Fill;
            lblAction.TextAlign = ContentAlignment.MiddleLeft;
            lblAction.Font = new Font(this.Font.FontFamily, 9F);
            layout.Controls.Add(lblAction, 0, 2);
            
            txtAction = new TextBox();
            txtAction.Dock = DockStyle.Fill;
            txtAction.Margin = new Padding(3, 5, 3, 3);
            txtAction.Font = new Font(this.Font.FontFamily, 10F);
            layout.Controls.Add(txtAction, 1, 2);
            formToolTip.SetToolTip(txtAction, "실행할 키보드 동작을 입력하세요. 오른쪽의 키 조합 패널을 이용해 쉽게 설정할 수 있습니다.");
            
            // 키 동작 예시 라벨
            Label lblActionInfo = new Label();
            lblActionInfo.Text = "예: CTRL+C, ALT+TAB, F5 등";
            lblActionInfo.Dock = DockStyle.Fill;
            lblActionInfo.TextAlign = ContentAlignment.TopLeft;
            lblActionInfo.ForeColor = Color.Gray;
            lblActionInfo.Font = new Font(this.Font.FontFamily, 8F, FontStyle.Italic);
            layout.Controls.Add(lblActionInfo, 1, 3);
            
            // 녹음 관련 컨트롤을 포함할 패널
            Panel recordingPanel = new Panel();
            recordingPanel.Dock = DockStyle.Fill;
            layout.Controls.Add(recordingPanel, 0, 4);
            layout.SetColumnSpan(recordingPanel, 2);
            
            // 녹음 버튼
            btnRecord = new Button();
            btnRecord.Text = "음성 녹음";
            btnRecord.Location = new Point(0, 10);
            btnRecord.Size = new Size(110, 35);
            btnRecord.Click += BtnRecord_Click;
            btnRecord.BackColor = Color.LightSkyBlue;
            btnRecord.FlatStyle = FlatStyle.Flat;
            btnRecord.Font = new Font(this.Font.FontFamily, 9F, FontStyle.Bold);
            recordingPanel.Controls.Add(btnRecord);
            formToolTip.SetToolTip(btnRecord, "클릭하여 음성을 녹음하고 자동으로 명령어를 인식합니다.");
            
            // 녹음 상태 표시줄
            progressRecording = new ProgressBar();
            progressRecording.Location = new Point(120, 10);
            progressRecording.Size = new Size(panel.Width - 150, 15);
            progressRecording.Minimum = 0;
            progressRecording.Maximum = 100;
            progressRecording.Value = 0;
            recordingPanel.Controls.Add(progressRecording);
            
            // 녹음 상태 라벨
            lblRecordingStatus = new Label();
            lblRecordingStatus.Text = "준비";
            lblRecordingStatus.Location = new Point(120, 30);
            lblRecordingStatus.Size = new Size(panel.Width - 150, 20);
            lblRecordingStatus.ForeColor = Color.Gray;
            lblRecordingStatus.Font = new Font(this.Font.FontFamily, 8F);
            recordingPanel.Controls.Add(lblRecordingStatus);
        }
        
        /// <summary>
        /// 오른쪽 상단 패널 - 키 조합 및 모디파이어 초기화
        /// </summary>
        private void InitializeKeyModifierPanel(GroupBox panel)
        {
            // 모디파이어 키 영역을 위한 패널
            Panel modifierPanel = new Panel();
            modifierPanel.Dock = DockStyle.Top;
            modifierPanel.Height = 70;
            modifierPanel.Padding = new Padding(5);
            panel.Controls.Add(modifierPanel);
            
            // 모디파이어 키 라벨
            Label lblModifiers = new Label();
            lblModifiers.Text = "모디파이어 키:";
            lblModifiers.Location = new Point(5, 10);
            lblModifiers.Size = new Size(100, 20);
            lblModifiers.Font = new Font(this.Font.FontFamily, 9F);
            modifierPanel.Controls.Add(lblModifiers);
            
            // 모디파이어 키 버튼들
            string[] modifierNames = { "Ctrl", "Alt", "Shift", "Win" };
            Color[] modifierColors = { Color.LightBlue, Color.LightGreen, Color.LightCoral, Color.LightGoldenrodYellow };
            int buttonWidth = 70;
            int gap = 10;
            int startX = 110;
            
            // Ctrl 버튼
            btnCtrl = new Button();
            btnCtrl.Text = modifierNames[0];
            btnCtrl.Location = new Point(startX, 5);
            btnCtrl.Size = new Size(buttonWidth, 60);
            btnCtrl.BackColor = Color.WhiteSmoke;
            btnCtrl.FlatStyle = FlatStyle.Flat;
            btnCtrl.Font = new Font(this.Font.FontFamily, 9F, FontStyle.Bold);
            btnCtrl.Click += (sender, e) => 
            {
                isCtrlPressed = !isCtrlPressed;
                btnCtrl.BackColor = isCtrlPressed ? modifierColors[0] : Color.WhiteSmoke;
                UpdateKeyAction();
                ShowHelp("Ctrl 키를 누른 상태로 다른 키를 누르는 조합입니다.");
            };
            modifierPanel.Controls.Add(btnCtrl);
            formToolTip.SetToolTip(btnCtrl, "Ctrl 키를 조합에 추가/제거합니다.");
            
            // Alt 버튼
            btnAlt = new Button();
            btnAlt.Text = modifierNames[1];
            btnAlt.Location = new Point(startX + buttonWidth + gap, 5);
            btnAlt.Size = new Size(buttonWidth, 60);
            btnAlt.BackColor = Color.WhiteSmoke;
            btnAlt.FlatStyle = FlatStyle.Flat;
            btnAlt.Font = new Font(this.Font.FontFamily, 9F, FontStyle.Bold);
            btnAlt.Click += (sender, e) => 
            {
                isAltPressed = !isAltPressed;
                btnAlt.BackColor = isAltPressed ? modifierColors[1] : Color.WhiteSmoke;
                UpdateKeyAction();
                ShowHelp("Alt 키를 누른 상태로 다른 키를 누르는 조합입니다.");
            };
            modifierPanel.Controls.Add(btnAlt);
            formToolTip.SetToolTip(btnAlt, "Alt 키를 조합에 추가/제거합니다.");
            
            // Shift 버튼
            btnShift = new Button();
            btnShift.Text = modifierNames[2];
            btnShift.Location = new Point(startX + (buttonWidth + gap) * 2, 5);
            btnShift.Size = new Size(buttonWidth, 60);
            btnShift.BackColor = Color.WhiteSmoke;
            btnShift.FlatStyle = FlatStyle.Flat;
            btnShift.Font = new Font(this.Font.FontFamily, 9F, FontStyle.Bold);
            btnShift.Click += (sender, e) => 
            {
                isShiftPressed = !isShiftPressed;
                btnShift.BackColor = isShiftPressed ? modifierColors[2] : Color.WhiteSmoke;
                UpdateKeyAction();
                ShowHelp("Shift 키를 누른 상태로 다른 키를 누르는 조합입니다.");
            };
            modifierPanel.Controls.Add(btnShift);
            formToolTip.SetToolTip(btnShift, "Shift 키를 조합에 추가/제거합니다.");
            
            // Win 버튼
            btnWin = new Button();
            btnWin.Text = modifierNames[3];
            btnWin.Location = new Point(startX + (buttonWidth + gap) * 3, 5);
            btnWin.Size = new Size(buttonWidth, 60);
            btnWin.BackColor = Color.WhiteSmoke;
            btnWin.FlatStyle = FlatStyle.Flat;
            btnWin.Font = new Font(this.Font.FontFamily, 9F, FontStyle.Bold);
            btnWin.Click += (sender, e) => 
            {
                isWinPressed = !isWinPressed;
                btnWin.BackColor = isWinPressed ? modifierColors[3] : Color.WhiteSmoke;
                UpdateKeyAction();
                ShowHelp("Windows 키를 누른 상태로 다른 키를 누르는 조합입니다.");
            };
            modifierPanel.Controls.Add(btnWin);
            formToolTip.SetToolTip(btnWin, "Windows 키를 조합에 추가/제거합니다.");
            
            // 가상 키패드를 위한 FlowLayoutPanel
            keyboardPanel = new FlowLayoutPanel();
            keyboardPanel.Dock = DockStyle.Fill;
            keyboardPanel.FlowDirection = FlowDirection.LeftToRight;
            keyboardPanel.WrapContents = true;
            keyboardPanel.AutoScroll = true;
            keyboardPanel.Padding = new Padding(5);
            panel.Controls.Add(keyboardPanel);
            
            // 기능 키 그룹 (F1-F12)
            AddKeyGroupHeader(keyboardPanel, "기능 키");
            for (int i = 1; i <= 12; i++)
            {
                string keyName = $"F{i}";
                AddKeyButton(keyboardPanel, keyName, $"F{i} 기능 키를 누릅니다.");
            }
            
            // 특수 키 그룹
            AddKeyGroupHeader(keyboardPanel, "특수 키");
            string[] specialKeys = new string[] 
            { 
                "Enter", "Esc", "Tab", "Space", "Backspace", 
                "Insert", "Delete", "Home", "End", "PageUp", "PageDown",
                "PrintScreen", "ScrollLock", "Pause"
            };
            
            foreach (string key in specialKeys)
            {
                AddKeyButton(keyboardPanel, key, $"{key} 키를 누릅니다.");
            }
            
            // 방향 키 그룹
            AddKeyGroupHeader(keyboardPanel, "방향 키");
            string[] directionKeys = new string[] { "Up", "Down", "Left", "Right" };
            foreach (string key in directionKeys)
            {
                AddKeyButton(keyboardPanel, key, $"{key} 방향 키를 누릅니다.");
            }
            
            // 숫자 키 그룹
            AddKeyGroupHeader(keyboardPanel, "숫자 키");
            for (int i = 0; i <= 9; i++)
            {
                AddKeyButton(keyboardPanel, i.ToString(), $"숫자 {i} 키를 누릅니다.");
            }
            
            // 알파벳 키 그룹
            AddKeyGroupHeader(keyboardPanel, "알파벳 키");
            for (char c = 'A'; c <= 'Z'; c++)
            {
                AddKeyButton(keyboardPanel, c.ToString(), $"{c} 키를 누릅니다.");
            }
        }
        
        /// <summary>
        /// 키패드에 키 그룹 헤더를 추가합니다.
        /// </summary>
        private void AddKeyGroupHeader(FlowLayoutPanel panel, string text)
        {
            Label header = new Label();
            header.Text = text;
            header.AutoSize = false;
            header.Width = panel.Width - 20;
            header.Height = 25;
            header.TextAlign = ContentAlignment.MiddleLeft;
            header.Font = new Font(this.Font.FontFamily, 9F, FontStyle.Bold);
            header.ForeColor = Color.DarkBlue;
            header.BorderStyle = BorderStyle.FixedSingle;
            header.BackColor = Color.LightCyan;
            header.Margin = new Padding(0, 10, 0, 5);
            panel.Controls.Add(header);
        }
        
        /// <summary>
        /// 키패드에 키 버튼을 추가합니다.
        /// </summary>
        private void AddKeyButton(FlowLayoutPanel panel, string keyText, string helpText)
        {
            if (panel == null) return;
            
            Button btn = new Button();
            btn.Text = keyText;
            btn.Size = new Size(60, 30);
            btn.Margin = new Padding(3);
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = Color.WhiteSmoke;
            btn.Click += (sender, e) => AddKeyToAction(keyText);
            btn.MouseEnter += (sender, e) => ShowHelp(helpText);
            
            if (formToolTip != null)
            {
                formToolTip.SetToolTip(btn, $"{keyText} 키를 키 동작에 추가합니다.");
            }
            
            panel.Controls.Add(btn);
        }
        
        /// <summary>
        /// 왼쪽 하단 패널 - 액션 타입 설정 초기화
        /// </summary>
        private void InitializeActionTypePanel(GroupBox panel)
        {
            // 패널 내 컨트롤 배치를 위한 테이블 레이아웃
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.RowCount = 3;
            layout.ColumnCount = 2;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            panel.Controls.Add(layout);
            
            // 액션 타입 라벨 및 콤보박스
            lblActionType = new Label();
            lblActionType.Text = "액션 타입:";
            lblActionType.Dock = DockStyle.Fill;
            lblActionType.TextAlign = ContentAlignment.MiddleLeft;
            lblActionType.Font = new Font(this.Font.FontFamily, 9F);
            layout.Controls.Add(lblActionType, 0, 0);
            
            cmbActionType = new ComboBox();
            cmbActionType.Dock = DockStyle.Fill;
            cmbActionType.Margin = new Padding(3, 5, 3, 3);
            cmbActionType.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbActionType.Font = new Font(this.Font.FontFamily, 10F);
            layout.Controls.Add(cmbActionType, 1, 0);
            formToolTip.SetToolTip(cmbActionType, "키 액션의 동작 방식을 선택합니다.");
            
            // 액션 타입 변경 이벤트 핸들러 추가는 InitializeActionTypes에서 수행
            
            // 액션 파라미터 라벨 및 입력 컨트롤
            lblActionParam = new Label();
            lblActionParam.Text = "파라미터:";
            lblActionParam.Dock = DockStyle.Fill;
            lblActionParam.TextAlign = ContentAlignment.MiddleLeft;
            lblActionParam.Font = new Font(this.Font.FontFamily, 9F);
            layout.Controls.Add(lblActionParam, 0, 1);
            
            // 파라미터 입력을 위한 패널 (수치 입력 + 설명 라벨)
            Panel paramPanel = new Panel();
            paramPanel.Dock = DockStyle.Fill;
            layout.Controls.Add(paramPanel, 1, 1);
            
            numActionParam = new NumericUpDown();
            numActionParam.Location = new Point(0, 5);
            numActionParam.Size = new Size(120, 25);
            numActionParam.Minimum = 0;
            numActionParam.Maximum = 10000;
            numActionParam.Font = new Font(this.Font.FontFamily, 10F);
            paramPanel.Controls.Add(numActionParam);
            formToolTip.SetToolTip(numActionParam, "선택한 액션 타입에 따른 파라미터 값을 설정합니다.");
            
            lblParamDescription = new Label();
            lblParamDescription.Text = "추가 정보 없음";
            lblParamDescription.Location = new Point(130, 8);
            lblParamDescription.Size = new Size(panel.Width - 160, 20);
            lblParamDescription.ForeColor = Color.Gray;
            lblParamDescription.Font = new Font(this.Font.FontFamily, 8F, FontStyle.Italic);
            paramPanel.Controls.Add(lblParamDescription);
            
            // 액션 타입별 상세 설명 영역
            RichTextBox txtActionTypeDescription = new RichTextBox();
            txtActionTypeDescription.Dock = DockStyle.Fill;
            txtActionTypeDescription.BackColor = Color.LightYellow;
            txtActionTypeDescription.ReadOnly = true;
            txtActionTypeDescription.BorderStyle = BorderStyle.None;
            txtActionTypeDescription.Font = new Font(this.Font.FontFamily, 9F);
            layout.Controls.Add(txtActionTypeDescription, 0, 2);
            layout.SetColumnSpan(txtActionTypeDescription, 2);
            
            // 액션 타입 변경 시 설명 업데이트
            cmbActionType.SelectedIndexChanged += (sender, e) => 
            {
                MacroActionType selectedType = (MacroActionType)cmbActionType.SelectedIndex;
                
                // 액션 타입별 상세 설명
                string description = GetActionTypeDescription(selectedType);
                txtActionTypeDescription.Text = description;
                
                // 도움말 영역에도 표시
                ShowHelp(description);
            };
        }
        
        /// <summary>
        /// 액션 타입별 상세 설명을 반환합니다.
        /// </summary>
        private string GetActionTypeDescription(MacroActionType actionType)
        {
            switch (actionType)
            {
                case MacroActionType.Default:
                    return "기본 키 액션: 지정된 키를 한 번 입력합니다.\r\n\r\n"
                         + "- 사용 예시: 단일 키 입력, 단축키 등\r\n"
                         + "- 필요한 파라미터: 없음";
                
                case MacroActionType.Toggle:
                    return "토글 키 액션: 키를 누르고 떼는 동작을 전환합니다.\r\n\r\n"
                         + "- 첫 번째 호출: 키를 누른 상태 유지\r\n"
                         + "- 두 번째 호출: 키를 뗌\r\n"
                         + "- 사용 예시: Caps Lock, Num Lock 등의 토글 키\r\n"
                         + "- 필요한 파라미터: 없음";
                
                case MacroActionType.Repeat:
                    return "반복 키 액션: 지정된 키를 여러 번 반복해서 입력합니다.\r\n\r\n"
                         + "- 사용 예시: 동일한 키를 여러 번 누르는 경우\r\n"
                         + "- 필요한 파라미터: 반복 횟수 (1~100회)";
                
                case MacroActionType.Hold:
                    return "홀드 키 액션: 지정된 키를 일정 시간 동안 누른 상태로 유지합니다.\r\n\r\n"
                         + "- 사용 예시: 게임에서 달리기, 아이템 사용 등\r\n"
                         + "- 필요한 파라미터: 유지 시간 (밀리초, 100~10000ms)";
                
                case MacroActionType.Turbo:
                    return "터보 키 액션: 지정된 키를 빠르게 연타합니다.\r\n\r\n"
                         + "- 사용 예시: 게임에서 빠른 공격, 연속 입력이 필요한 경우\r\n"
                         + "- 필요한 파라미터: 입력 간격 (밀리초, 10~500ms)\r\n"
                         + "- 기본 연타 횟수: 10회";
                
                case MacroActionType.Combo:
                    return "콤보 키 액션: 여러 키를 순차적으로 입력합니다.\r\n\r\n"
                         + "- 사용 예시: 게임 콤보 기술, 여러 단계의 단축키\r\n"
                         + "- 필요한 파라미터: 키 사이 간격 (밀리초, 10~1000ms)\r\n"
                         + "- 키 입력 형식: 쉼표(,)로 구분하여 입력";
                
                default:
                    return "선택한 액션 타입에 대한 정보가 없습니다.";
            }
        }
        
        /// <summary>
        /// 오른쪽 하단 패널 - 도움말 및 버튼 초기화
        /// </summary>
        private void InitializeHelpAndButtonPanel(GroupBox panel)
        {
            // 패널 내 레이아웃
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.RowCount = 2;
            layout.ColumnCount = 1;
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 70F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 30F));
            panel.Controls.Add(layout);
            
            // 도움말 텍스트 영역
            txtHelpInfo = new RichTextBox();
            txtHelpInfo.Dock = DockStyle.Fill;
            txtHelpInfo.BackColor = Color.LightYellow;
            txtHelpInfo.ReadOnly = true;
            txtHelpInfo.BorderStyle = BorderStyle.None;
            txtHelpInfo.Font = new Font(this.Font.FontFamily, 9F);
            txtHelpInfo.Text = "마우스를 컨트롤 위에 올리면 해당 기능에 대한 도움말이 여기에 표시됩니다.";
            layout.Controls.Add(txtHelpInfo, 0, 0);
            
            // 확인/취소 버튼 패널
            Panel buttonPanel = new Panel();
            buttonPanel.Dock = DockStyle.Fill;
            layout.Controls.Add(buttonPanel, 0, 1);
            
            Button btnOk = new Button();
            btnOk.Text = "확인";
            btnOk.Location = new Point(buttonPanel.Width / 2 - 170, 20);
            btnOk.Size = new Size(150, 40);
            btnOk.BackColor = Color.LightGreen;
            btnOk.FlatStyle = FlatStyle.Flat;
            btnOk.Font = new Font(this.Font.FontFamily, 10F, FontStyle.Bold);
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
                SelectedActionType = (MacroActionType)cmbActionType.SelectedIndex;
                SelectedActionParam = (int)numActionParam.Value;
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            buttonPanel.Controls.Add(btnOk);
            formToolTip.SetToolTip(btnOk, "변경 사항을 저장하고 창을 닫습니다.");
            
            Button btnCancel = new Button();
            btnCancel.Text = "취소";
            btnCancel.Location = new Point(buttonPanel.Width / 2 + 20, 20);
            btnCancel.Size = new Size(150, 40);
            btnCancel.BackColor = Color.LightCoral;
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.Font = new Font(this.Font.FontFamily, 10F, FontStyle.Bold);
            btnCancel.Click += (sender, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };
            buttonPanel.Controls.Add(btnCancel);
            formToolTip.SetToolTip(btnCancel, "변경 사항을 취소하고 창을 닫습니다.");
            
            // 버튼 위치 조정을 위한 이벤트
            buttonPanel.Resize += (sender, e) => 
            {
                btnOk.Location = new Point(buttonPanel.Width / 2 - 170, 20);
                btnCancel.Location = new Point(buttonPanel.Width / 2 + 20, 20);
            };
        }
        
        /// <summary>
        /// 도움말 정보를 표시합니다.
        /// </summary>
        private void ShowHelp(string helpText)
        {
            if (txtHelpInfo != null && !string.IsNullOrEmpty(helpText))
            {
                txtHelpInfo.Text = helpText;
            }
        }

        /// <summary>
        /// 키 액션 텍스트상자에 키를 추가합니다.
        /// </summary>
        private void AddKeyToAction(string key)
        {
            // 기존에 다른 키가 있으면 + 기호로 연결, 아니면 새로 설정
            if (!string.IsNullOrEmpty(txtAction.Text) && !txtAction.Text.EndsWith("+"))
            {
                // 콤보 타입이면 쉼표로 구분
                if ((MacroActionType)cmbActionType.SelectedIndex == MacroActionType.Combo)
                {
                    txtAction.Text += ", " + key;
                }
                else
                {
                    // 이미 모디파이어 키가 선택되어 있는 상태면 "+"로 연결
                    txtAction.Text += "+" + key;
                }
            }
            else
            {
                // 비어있거나 +로 끝나면 바로 추가
                txtAction.Text += key;
            }
        }
        
        /// <summary>
        /// 선택된 모디파이어 키에 따라 키 액션을 업데이트합니다.
        /// </summary>
        private void UpdateKeyAction()
        {
            // 기존 키 액션에서 모디파이어 부분을 제거
            string originalAction = txtAction.Text;
            string baseKey = originalAction;
            
            // 기존 모디파이어 키를 찾아 제거
            if (originalAction.Contains("+"))
            {
                string[] parts = originalAction.Split('+');
                baseKey = parts[parts.Length - 1];
                
                // 모든 모디파이어 제거
                foreach (string modifier in new[] { "CTRL", "ALT", "SHIFT", "WIN" })
                {
                    baseKey = baseKey.Replace(modifier + "+", "");
                }
            }
            
            // 새 모디파이어 키 조합 구성
            string newAction = "";
            if (isCtrlPressed) newAction += "CTRL+";
            if (isAltPressed) newAction += "ALT+";
            if (isShiftPressed) newAction += "SHIFT+";
            if (isWinPressed) newAction += "WIN+";
            
            // 기본 키가 있으면 추가
            if (!string.IsNullOrEmpty(baseKey))
            {
                newAction += baseKey;
            }
            
            txtAction.Text = newAction;
        }
        
        /// <summary>
        /// 액션 타입 관련 컨트롤을 초기화합니다.
        /// </summary>
        private void InitializeActionTypes()
        {
            // cmbActionType이 null인지 확인
            if (cmbActionType == null) return;
            
            // 콤보박스에 액션 타입 목록 추가
            cmbActionType.Items.Clear();
            cmbActionType.Items.Add("기본 (한 번 입력)");
            cmbActionType.Items.Add("토글 (키 누르기/떼기 전환)");
            cmbActionType.Items.Add("반복 (여러 번 입력)");
            cmbActionType.Items.Add("홀드 (키 누르고 유지)");
            cmbActionType.Items.Add("터보 (빠른 키 연타)");
            cmbActionType.Items.Add("콤보 (키 순차 입력)");
            
            // 기본 선택
            cmbActionType.SelectedIndex = (int)SelectedActionType;
            
            // 액션 타입 변경 이벤트 핸들러
            cmbActionType.SelectedIndexChanged += (sender, e) => 
            {
                if (cmbActionType == null) return;
                
                MacroActionType selectedType = (MacroActionType)cmbActionType.SelectedIndex;
                UpdateActionParamControl(selectedType);
            };
            
            // 초기 파라미터 컨트롤 상태 설정
            UpdateActionParamControl((MacroActionType)cmbActionType.SelectedIndex);
        }
        
        /// <summary>
        /// 선택된 액션 타입에 따라 파라미터 컨트롤을 업데이트합니다.
        /// </summary>
        private void UpdateActionParamControl(MacroActionType actionType)
        {
            // numActionParam이나 lblParamDescription이 null인지 확인
            if (numActionParam == null || lblParamDescription == null) return;
            
            // 액션 타입에 따라 파라미터 설정 조정
            switch (actionType)
            {
                case MacroActionType.Default:
                    numActionParam.Enabled = false;
                    numActionParam.Value = 0;
                    lblParamDescription.Text = "파라미터 필요 없음";
                    break;
                
                case MacroActionType.Toggle:
                    numActionParam.Enabled = false;
                    numActionParam.Value = 0;
                    lblParamDescription.Text = "파라미터 필요 없음";
                    break;
                
                case MacroActionType.Repeat:
                    numActionParam.Enabled = true;
                    numActionParam.Minimum = 1;
                    numActionParam.Maximum = 100;
                    numActionParam.Value = Math.Max(1, Math.Min(100, SelectedActionParam > 0 ? SelectedActionParam : 3));
                    numActionParam.Increment = 1;
                    lblParamDescription.Text = "반복 횟수";
                    break;
                
                case MacroActionType.Hold:
                    numActionParam.Enabled = true;
                    numActionParam.Minimum = 100;
                    numActionParam.Maximum = 10000;
                    numActionParam.Value = Math.Max(100, Math.Min(10000, SelectedActionParam > 0 ? SelectedActionParam : 1000));
                    numActionParam.Increment = 100;
                    lblParamDescription.Text = "유지 시간 (밀리초)";
                    break;
                
                case MacroActionType.Turbo:
                    numActionParam.Enabled = true;
                    numActionParam.Minimum = 10;
                    numActionParam.Maximum = 500;
                    numActionParam.Value = Math.Max(10, Math.Min(500, SelectedActionParam > 0 ? SelectedActionParam : 50));
                    numActionParam.Increment = 10;
                    lblParamDescription.Text = "입력 간격 (밀리초)";
                    break;
                
                case MacroActionType.Combo:
                    numActionParam.Enabled = true;
                    numActionParam.Minimum = 10;
                    numActionParam.Maximum = 1000;
                    numActionParam.Value = Math.Max(10, Math.Min(1000, SelectedActionParam > 0 ? SelectedActionParam : 100));
                    numActionParam.Increment = 10;
                    lblParamDescription.Text = "키 간격 (밀리초)";
                    break;
            }
        }
        
        /// <summary>
        /// 키 액션 문자열을 파싱하여 UI 상태를 업데이트합니다.
        /// </summary>
        private void ParseKeyAction(string keyAction)
        {
            if (string.IsNullOrEmpty(keyAction)) return;
            
            // 모디파이어 키 상태 초기화
            isCtrlPressed = keyAction.Contains("CTRL+");
            isAltPressed = keyAction.Contains("ALT+");
            isShiftPressed = keyAction.Contains("SHIFT+");
            isWinPressed = keyAction.Contains("WIN+");
            
            // 버튼 상태 업데이트 - null 체크 추가
            if (btnCtrl != null) btnCtrl.BackColor = isCtrlPressed ? Color.LightBlue : Color.WhiteSmoke;
            if (btnAlt != null) btnAlt.BackColor = isAltPressed ? Color.LightGreen : Color.WhiteSmoke;
            if (btnShift != null) btnShift.BackColor = isShiftPressed ? Color.LightCoral : Color.WhiteSmoke;
            if (btnWin != null) btnWin.BackColor = isWinPressed ? Color.LightGoldenrodYellow : Color.WhiteSmoke;
        }
        
        /// <summary>
        /// 키 버튼 설정과 관련된 초기화를 수행합니다.
        /// </summary>
        private void SetupKeyButtons()
        {
            // 이 메서드는 더 이상 사용하지 않음 - 새 UI 레이아웃에서는 InitializeKeyModifierPanel에서 처리
            // 단, 기존 코드와의 호환성을 위해 빈 메서드로 유지
        }
        
        /// <summary>
        /// 일반 키 메뉴를 표시합니다.
        /// </summary>
        private void ShowCommonKeysMenu()
        {
            // 이 메서드는 더 이상 사용하지 않음 - 새 UI 레이아웃에서는 키 버튼이 직접 FlowLayoutPanel에 추가됨
            // 단, 기존 코드와의 호환성을 위해 빈 메서드로 유지
        }
    }
}
