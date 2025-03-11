using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Media;
using VoiceMacro.Services;

namespace VoiceMacro
{
    /// <summary>
    /// 색상을 변경할 수 있는 커스텀 ProgressBar
    /// </summary>
    public class ColorProgressBar : ProgressBar
    {
        private Color barColor = Color.Green;

        public Color BarColor
        {
            get { return barColor; }
            set 
            { 
                barColor = value;
                Invalidate(); // 색상 변경 시 다시 그림
            }
        }

        public ColorProgressBar()
        {
            this.SetStyle(ControlStyles.UserPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle rec = new Rectangle(0, 0, this.Width, this.Height);
            
            if (ProgressBarRenderer.IsSupported)
                ProgressBarRenderer.DrawHorizontalBar(e.Graphics, rec);
            
            rec.Width = (int)(rec.Width * ((double)Value / Maximum)) - 4;
            rec.Height -= 4;
            rec.X += 2;
            rec.Y += 2;
            
            // 현재 설정된 색상으로 채움
            using (SolidBrush brush = new SolidBrush(barColor))
            {
                e.Graphics.FillRectangle(brush, rec);
            }
        }
    }

    public partial class MainForm : Form
    {
        private NotifyIcon trayIcon;
        private VoiceRecognitionService voiceRecognizer;
        private MacroService macroService;
        private bool isListening = false;
        private RichTextBox rtbLog;
        private Button btnClearLog;
        private CheckBox chkAutoScroll;
        private CheckBox chkDetailedLog;
        private Button chkPlayBeep;
        private bool showDetailedLog = false;
        private bool playBeepSound = true;
        private Button btnStartStop;
        private Button btnAddMacro;
        private Button btnRemoveMacro;
        private Button btnSettings;
        private Button btnPresets;
        private Button btnCopyMacro;
        private Button btnEditMacro;
        private ListView lstMacros;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;
        private ToolTip toolTip;
        private TrackBar tbarMicVolume;      // 마이크 볼륨 조절 슬라이더
        private ColorProgressBar pbarMicLevel;    // 마이크 레벨 표시 프로그레스바
        private Label lblMicVolume;          // 마이크 볼륨 레이블

        public MainForm()
        {
            InitializeComponent();
            InitializeTrayIcon();
            InitializeServices();

            // MacroService 이벤트 연결
            macroService.MacroExecuted += MacroService_MacroExecuted;
            macroService.StatusChanged += MacroService_StatusChanged;

            // VoiceRecognition 이벤트 연결
            voiceRecognizer.SpeechRecognized += VoiceRecognizer_SpeechRecognized;
            voiceRecognizer.StatusChanged += VoiceRecognizer_StatusChanged;
            voiceRecognizer.AudioLevelChanged += VoiceRecognizer_AudioLevelChanged;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // 툴크 초기화
            this.toolTip = new ToolTip();
            this.toolTip.AutoPopDelay = 5000;
            this.toolTip.InitialDelay = 1000;
            this.toolTip.ReshowDelay = 500;
            
            // MainForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1000, 700);
            this.MinimumSize = new Size(900, 650);
            this.Name = "MainForm";
            this.Text = "음성 매크로";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Resize += new System.EventHandler(this.MainForm_Resize);
            
            // 상단 컨트롤 패널 (고정 높이)
            Panel topPanel = new Panel();
            topPanel.Dock = DockStyle.Top;
            topPanel.Height = 80;
            topPanel.BackColor = Color.FromArgb(240, 240, 240);
            topPanel.Padding = new Padding(10);
            
            // 컨트롤 그룹 (상단 패널)
            GroupBox controlGroup = new GroupBox();
            controlGroup.Text = "컨트롤";
            controlGroup.Location = new Point(10, 5);
            controlGroup.Size = new Size(650, 70);
            controlGroup.Padding = new Padding(5);
            
            // 시작/중지 버튼
            this.btnStartStop = new Button();
            this.btnStartStop.Location = new Point(20, 25);
            this.btnStartStop.Size = new Size(120, 30);
            this.btnStartStop.Text = "시작";
            this.btnStartStop.BackColor = Color.LightGreen;
            this.btnStartStop.Font = new Font(this.Font, FontStyle.Bold);
            this.btnStartStop.Click += new EventHandler(this.btnStartStop_Click);
            controlGroup.Controls.Add(this.btnStartStop);
            
            // 설정 버튼
            this.btnSettings = new Button();
            this.btnSettings.Location = new Point(160, 25);
            this.btnSettings.Size = new Size(120, 30);
            this.btnSettings.Text = "설정";
            this.btnSettings.Click += new EventHandler(this.btnSettings_Click);
            controlGroup.Controls.Add(this.btnSettings);
            
            // 프리셋 버튼
            this.btnPresets = new Button();
            this.btnPresets.Location = new Point(300, 25);
            this.btnPresets.Size = new Size(120, 30);
            this.btnPresets.Text = "프리셋";
            this.btnPresets.Click += new EventHandler(this.btnPresets_Click);
            controlGroup.Controls.Add(this.btnPresets);
            
            // 알림음 버튼 추가
            this.chkPlayBeep = new Button();
            this.chkPlayBeep.Location = new Point(440, 25);
            this.chkPlayBeep.Size = new Size(120, 30);
            this.chkPlayBeep.Text = "알림음: 켜짐";
            this.chkPlayBeep.BackColor = Color.LightSkyBlue;
            this.chkPlayBeep.Click += new EventHandler(this.ChkPlayBeep_Click);
            this.toolTip.SetToolTip(this.chkPlayBeep, "시스템 트레이에서 작동 중일 때 음성인식 성공 시 알림음 재생");
            controlGroup.Controls.Add(this.chkPlayBeep);
            
            topPanel.Controls.Add(controlGroup);
            
            // 메인 콘텐츠 패널 (상단과 하단 사이)
            Panel contentPanel = new Panel();
            contentPanel.Dock = DockStyle.Fill;
            contentPanel.Padding = new Padding(10);
            
            // 좌측 패널 생성 (로그 영역)
            Panel leftPanel = new Panel();
            leftPanel.Dock = DockStyle.Left;
            leftPanel.Width = 300;
            leftPanel.Padding = new Padding(0, 0, 10, 0); // 오른쪽에 패딩 추가
            
            // 로그 그룹
            GroupBox logGroup = new GroupBox();
            logGroup.Text = "로그";
            logGroup.Dock = DockStyle.Fill;
            logGroup.Padding = new Padding(10);
            leftPanel.Controls.Add(logGroup);
            
            // 로그 텍스트 박스
            this.rtbLog = new RichTextBox();
            this.rtbLog.Dock = DockStyle.Top;
            this.rtbLog.Height = 520;
            this.rtbLog.ReadOnly = true;
            this.rtbLog.BackColor = Color.White;
            this.rtbLog.Font = new Font("Consolas", 9F);
            this.rtbLog.ScrollBars = RichTextBoxScrollBars.Vertical;
            logGroup.Controls.Add(this.rtbLog);
            
            // 로그 컨트롤 패널
            Panel logControlPanel = new Panel();
            logControlPanel.Dock = DockStyle.Bottom;
            logControlPanel.Height = 40;
            logGroup.Controls.Add(logControlPanel);
            
            // 로그 컨트롤 버튼 및 체크박스
            this.btnClearLog = new Button();
            this.btnClearLog.Text = "로그 지우기";
            this.btnClearLog.Location = new Point(5, 5);
            this.btnClearLog.Size = new Size(80, 25);
            this.btnClearLog.Click += new EventHandler(this.BtnClearLog_Click);
            logControlPanel.Controls.Add(this.btnClearLog);
            
            // 자동 스크롤 체크박스
            this.chkAutoScroll = new CheckBox();
            this.chkAutoScroll.Text = "자동 스크롤";
            this.chkAutoScroll.Location = new Point(90, 8);
            this.chkAutoScroll.Size = new Size(85, 20);
            this.chkAutoScroll.Checked = true;
            logControlPanel.Controls.Add(this.chkAutoScroll);
            
            // 상세 로그 체크박스
            this.chkDetailedLog = new CheckBox();
            this.chkDetailedLog.Text = "상세 로그";
            this.chkDetailedLog.Location = new Point(180, 8);
            this.chkDetailedLog.Size = new Size(75, 20);
            this.chkDetailedLog.CheckedChanged += new EventHandler(this.ChkDetailedLog_CheckedChanged);
            logControlPanel.Controls.Add(this.chkDetailedLog);
           
            
            // 오른쪽 패널 (매크로 관리 영역)
            Panel rightPanel = new Panel();
            rightPanel.Dock = DockStyle.Fill;
            rightPanel.Padding = new Padding(0, 0, 0, 0);
            
            // 매크로 관리 그룹
            GroupBox macroManageGroup = new GroupBox();
            macroManageGroup.Text = "매크로 관리";
            macroManageGroup.Dock = DockStyle.Fill;
            macroManageGroup.Padding = new Padding(10);
            rightPanel.Controls.Add(macroManageGroup);
            
            // 마이크 컨트롤 패널 (NEW)
            Panel micControlPanel = new Panel();
            micControlPanel.Dock = DockStyle.Top;
            micControlPanel.Height = 40;
            micControlPanel.BackColor = Color.FromArgb(245, 245, 245); // 약간 회색 배경으로 구분
            
            // 마이크 볼륨 레이블
            this.lblMicVolume = new Label();
            this.lblMicVolume.Text = "마이크 볼륨:";
            this.lblMicVolume.Location = new Point(10, 12);
            this.lblMicVolume.Size = new Size(80, 20);
            this.lblMicVolume.TextAlign = ContentAlignment.MiddleRight;
            micControlPanel.Controls.Add(this.lblMicVolume);
            
            // 마이크 볼륨 슬라이더
            this.tbarMicVolume = new TrackBar();
            this.tbarMicVolume.Location = new Point(95, 5);
            this.tbarMicVolume.Size = new Size(150, 30);
            this.tbarMicVolume.Minimum = 0;
            this.tbarMicVolume.Maximum = 100;
            this.tbarMicVolume.Value = 100; // 기본값은 100%
            this.tbarMicVolume.TickFrequency = 10;
            this.tbarMicVolume.SmallChange = 1;
            this.tbarMicVolume.LargeChange = 10;
            this.tbarMicVolume.Orientation = Orientation.Horizontal;
            this.tbarMicVolume.ValueChanged += new EventHandler(this.TbarMicVolume_ValueChanged);
            this.toolTip.SetToolTip(this.tbarMicVolume, "마이크 볼륨을 조절합니다.");
            micControlPanel.Controls.Add(this.tbarMicVolume);
            
            // 마이크 레벨 미터
            this.pbarMicLevel = new ColorProgressBar();
            this.pbarMicLevel.Location = new Point(395, 12);
            this.pbarMicLevel.Size = new Size(130, 15);
            this.pbarMicLevel.Minimum = 0;
            this.pbarMicLevel.Maximum = 100;
            this.toolTip.SetToolTip(this.pbarMicLevel, "현재 마이크 입력 레벨");
            micControlPanel.Controls.Add(this.pbarMicLevel);
            
            // 레벨 레이블 추가
            Label lblMicLevel = new Label();
            lblMicLevel.Text = "입력 레벨:";
            lblMicLevel.Location = new Point(385, 12);
            lblMicLevel.Size = new Size(130, 20);
            lblMicLevel.TextAlign = ContentAlignment.MiddleLeft;
            lblMicLevel.ForeColor = Color.Gray;
            micControlPanel.Controls.Add(lblMicLevel);
            
            // 매크로 버튼 패널
            Panel macroButtonPanel = new Panel();
            macroButtonPanel.Dock = DockStyle.Top;
            macroButtonPanel.Height = 40;
            
            // 매크로 추가 버튼
            this.btnAddMacro = new Button();
            this.btnAddMacro.Location = new Point(5, 5);
            this.btnAddMacro.Size = new Size(120, 30);
            this.btnAddMacro.Text = "매크로 추가";
            this.btnAddMacro.BackColor = Color.LightBlue;
            this.btnAddMacro.Click += new EventHandler(this.btnAddMacro_Click);
            macroButtonPanel.Controls.Add(this.btnAddMacro);
            
            // 매크로 삭제 버튼
            this.btnRemoveMacro = new Button();
            this.btnRemoveMacro.Location = new Point(135, 5);
            this.btnRemoveMacro.Size = new Size(120, 30);
            this.btnRemoveMacro.Text = "매크로 삭제";
            this.btnRemoveMacro.Click += new EventHandler(this.btnRemoveMacro_Click);
            macroButtonPanel.Controls.Add(this.btnRemoveMacro);
            
            // 매크로 복사 버튼
            this.btnCopyMacro = new Button();
            this.btnCopyMacro.Location = new Point(265, 5);
            this.btnCopyMacro.Size = new Size(120, 30);
            this.btnCopyMacro.Text = "매크로 복사";
            this.btnCopyMacro.Click += new EventHandler(this.btnCopyMacro_Click);
            macroButtonPanel.Controls.Add(this.btnCopyMacro);
            
            // 매크로 수정 버튼
            this.btnEditMacro = new Button();
            this.btnEditMacro.Location = new Point(395, 5);
            this.btnEditMacro.Size = new Size(120, 30);
            this.btnEditMacro.Text = "매크로 수정";
            this.btnEditMacro.Click += new EventHandler(this.btnEditMacro_Click);
            macroButtonPanel.Controls.Add(this.btnEditMacro);
            
            // 10포인트 상단 공간을 위한 패널 추가
            Panel macroTopSpacePanel = new Panel();
            macroTopSpacePanel.Dock = DockStyle.Top;
            macroTopSpacePanel.Height = 10;
            
            // 매크로 목록
            this.lstMacros = new ListView();
            this.lstMacros.Dock = DockStyle.Fill;
            this.lstMacros.View = View.Details;
            this.lstMacros.FullRowSelect = true;
            this.lstMacros.HideSelection = false;
            this.lstMacros.GridLines = true;
            this.lstMacros.BackColor = Color.White;
            
            // 열 추가
            this.lstMacros.Columns.Add("키워드", 150);
            this.lstMacros.Columns.Add("키 동작", 150);
            this.lstMacros.Columns.Add("액션 타입", 100);
            this.lstMacros.Columns.Add("파라미터", 150);
            
            // 컨트롤을 추가하는 순서가 중요합니다 - Fill이 가장 먼저, 그 다음 순서대로 추가해야합니다
            macroManageGroup.Controls.Add(this.lstMacros);      // Fill - 가장 먼저 추가
            macroManageGroup.Controls.Add(macroTopSpacePanel);  // Top - 그 다음 추가
            macroManageGroup.Controls.Add(macroButtonPanel);    // Top - 그 다음 추가
            macroManageGroup.Controls.Add(micControlPanel);     // Top - 가장 위에 추가
            
            // 상태 표시줄
            this.statusStrip = new StatusStrip();
            this.statusLabel = new ToolStripStatusLabel();
            this.statusLabel.Text = "준비";
            this.statusStrip.Items.Add(this.statusLabel);
            
            // 패널들을 컨텐츠 패널에 추가 (순서 중요)
            contentPanel.Controls.Add(rightPanel); // 먼저 rightPanel을 추가 (Fill로 설정되어 있음)
            contentPanel.Controls.Add(leftPanel);  // 그 다음 leftPanel을 추가 (Left에 고정됨)
            
            // 폼에 컨트롤 추가 (순서 중요)
            this.Controls.Add(this.statusStrip);   // 상태바 (항상 가장 아래)
            this.Controls.Add(contentPanel);       // 컨텐츠 패널 (Fill 공간 채움)
            this.Controls.Add(topPanel);           // 상단 패널 (항상 위에)
            
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void BtnClearLog_Click(object sender, EventArgs e)
        {
            rtbLog.Clear();
            AddLogMessage("로그가 지워졌습니다.", LogMessageType.Info);
        }

        private void VoiceRecognizer_StatusChanged(object? sender, string status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => VoiceRecognizer_StatusChanged(sender, status)));
                return;
            }

            statusLabel.Text = status;
            
            // 중요한 상태 메시지이거나 상세 로그가 활성화된 경우에만 로그에 추가
            bool isImportantStatus = 
                status.Contains("초기화") || 
                status.Contains("시작") || 
                status.Contains("중지") || 
                status.Contains("오류") || 
                status.Contains("실패") ||
                status.Contains("마이크");
                
            if (isImportantStatus || showDetailedLog)
            {
                AddLogMessage($"상태: {status}", LogMessageType.Info);
            }
        }

        private void VoiceRecognizer_SpeechRecognized(object? sender, string text)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => VoiceRecognizer_SpeechRecognized(sender, text)));
                return;
            }

            AddLogMessage($"인식됨: {text}", LogMessageType.Recognition);
            
            // 프로그램이 트레이에 있거나 최소화되어 있을 때 비프음 재생
            if (WindowState == FormWindowState.Minimized || !this.Visible)
            {
                PlayBeep();
            }
        }

        private void MacroService_StatusChanged(object? sender, string status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => MacroService_StatusChanged(sender, status)));
                return;
            }

            // 중요한 매크로 상태 메시지이거나 상세 로그가 활성화된 경우에만 로그에 추가
            bool isImportantStatus = 
                status.Contains("오류") || 
                status.Contains("실패") || 
                status.Contains("성공") || 
                status.Contains("완료") || 
                status.Contains("일치하는 매크로 없음");
                
            if (isImportantStatus || showDetailedLog)
            {
                AddLogMessage($"매크로: {status}", LogMessageType.Info);
            }
        }

        private void MacroService_MacroExecuted(object? sender, MacroExecutionEventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => MacroService_MacroExecuted(sender, e)));
                return;
            }

            if (e.Success)
            {
                AddLogMessage($"성공: '{e.Keyword}' → {e.KeyAction}", LogMessageType.Success);
                
                // 프로그램이 트레이에 있거나 최소화되어 있을 때 비프음 재생
                if (WindowState == FormWindowState.Minimized || !this.Visible)
                {
                    PlayBeep();
                }
            }
            else
            {
                AddLogMessage($"실패: '{e.Keyword}' → {e.KeyAction} ({e.ErrorMessage})", LogMessageType.Error);
            }
        }

        /// <summary>
        /// 비프음을 재생합니다.
        /// </summary>
        private void PlayBeep()
        {
            try
            {
                // 비프음 재생이 활성화된 경우에만 소리 재생
                if (playBeepSound)
                {
                    // 시스템 알림음 (별표 소리) 재생
                    SystemSounds.Asterisk.Play();
                    
                    // 디버깅을 위한 로그 메시지 추가
                    if (showDetailedLog)
                    {
                        AddLogMessage("비프음이 재생되었습니다.", LogMessageType.Info);
                    }
                }
            }
            catch (Exception ex)
            {
                AddLogMessage($"비프음 재생 오류: {ex.Message}", LogMessageType.Error);
            }
        }

        /// <summary>
        /// 로그 메시지 유형
        /// </summary>
        private enum LogMessageType
        {
            Info,       // 일반 정보
            Success,    // 성공
            Error,      // 오류
            Recognition // 음성 인식
        }

        /// <summary>
        /// 로그 메시지를 추가합니다.
        /// </summary>
        /// <param name="message">표시할 메시지</param>
        /// <param name="type">메시지 유형</param>
        private void AddLogMessage(string message, LogMessageType type)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => AddLogMessage(message, type)));
                return;
            }

            // 시간 정보 추가
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string formattedMessage = $"[{timestamp}] {message}{Environment.NewLine}";

            // 메시지 유형에 따른 색상 설정
            Color textColor;
            switch (type)
            {
                case LogMessageType.Success:
                    textColor = Color.Green;
                    break;
                case LogMessageType.Error:
                    textColor = Color.Red;
                    break;
                case LogMessageType.Recognition:
                    textColor = Color.Blue;
                    break;
                default:
                    textColor = Color.Black;
                    break;
            }

            // 메시지 추가
            rtbLog.SelectionStart = rtbLog.TextLength;
            rtbLog.SelectionLength = 0;
            rtbLog.SelectionColor = textColor;
            rtbLog.AppendText(formattedMessage);
            rtbLog.SelectionColor = rtbLog.ForeColor;

            // 자동 스크롤 활성화되어 있으면 스크롤 아래로 이동
            if (chkAutoScroll.Checked)
            {
                rtbLog.SelectionStart = rtbLog.Text.Length;
                rtbLog.ScrollToCaret();
            }

            // 로그 크기 제한 (5000줄 초과 시 오래된 로그 삭제)
            const int MaxLogLines = 5000;
            if (rtbLog.Lines.Length > MaxLogLines)
            {
                int linesToRemove = rtbLog.Lines.Length - MaxLogLines;
                int charsToRemove = 0;
                
                for (int i = 0; i < linesToRemove; i++)
                {
                    charsToRemove += rtbLog.Lines[i].Length + Environment.NewLine.Length;
                }
                
                rtbLog.Select(0, charsToRemove);
                rtbLog.SelectedText = "";
            }
        }

        private void InitializeTrayIcon()
        {
            trayIcon = new NotifyIcon();
            trayIcon.Icon = SystemIcons.Application;
            trayIcon.Text = "음성 매크로";
            trayIcon.Visible = true;

            ContextMenuStrip trayMenu = new ContextMenuStrip();
            
            ToolStripMenuItem showItem = new ToolStripMenuItem("보기");
            showItem.Click += (sender, e) => { 
                this.Show(); 
                this.WindowState = FormWindowState.Normal; 
                this.ShowInTaskbar = true;
                this.Activate();
            };
            trayMenu.Items.Add(showItem);

            ToolStripMenuItem startStopItem = new ToolStripMenuItem("시작/정지");
            startStopItem.Click += (sender, e) => { ToggleListening(); };
            trayMenu.Items.Add(startStopItem);

            ToolStripMenuItem settingsItem = new ToolStripMenuItem("설정");
            settingsItem.Click += (sender, e) => 
            { 
                // 설정 대화상자 표시
                this.Show();
                this.WindowState = FormWindowState.Normal;
                btnSettings_Click(sender, e);
            };
            trayMenu.Items.Add(settingsItem);

            ToolStripMenuItem presetsItem = new ToolStripMenuItem("프리셋 관리");
            presetsItem.Click += (sender, e) => 
            { 
                // 프리셋 관리 대화상자 표시
                this.Show();
                this.WindowState = FormWindowState.Normal;
                btnPresets_Click(sender, e);
            };
            trayMenu.Items.Add(presetsItem);

            trayMenu.Items.Add(new ToolStripSeparator());

            ToolStripMenuItem exitItem = new ToolStripMenuItem("종료");
            exitItem.Click += (sender, e) => { 
                try
                {
                    // 프로그램 종료 시 리소스 정리
                    // 리스닝 중이라면 먼저 중지
                    if (isListening && voiceRecognizer != null)
                    {
                        isListening = false;
                        try
                        {
                            voiceRecognizer.StopListening();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"음성인식 중지 중 오류: {ex.Message}");
                        }
                    }
                    
                    // 그 다음 리소스 정리
                    try 
                    {
                        if (voiceRecognizer != null)
                        {
                            voiceRecognizer.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"음성인식 자원 해제 중 오류: {ex.Message}");
                    }
                    
                    // 트레이 아이콘 정리
                    if (trayIcon != null)
                    {
                        try
                        {
                            trayIcon.Visible = false;
                            trayIcon.Dispose();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"트레이 아이콘 해제 중 오류: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 종료 시 예외가 발생해도 사용자에게는 보여주지 않고 계속 종료 진행
                    System.Diagnostics.Debug.WriteLine($"종료 중 오류 발생: {ex.Message}");
                }
                finally
                {
                    // 프로그램 종료
                    Application.Exit();
                }
            };
            trayMenu.Items.Add(exitItem);

            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.MouseDoubleClick += (sender, e) => 
            { 
                if (e.Button == MouseButtons.Left)
                {
                    this.Show();
                    this.WindowState = FormWindowState.Normal;
                    this.ShowInTaskbar = true;
                    this.Activate(); // 창을 활성화
                }
            };
        }

        private void InitializeServices()
        {
            macroService = new MacroService();
            voiceRecognizer = new VoiceRecognitionService(macroService);
            
            // 마이크 레벨 이벤트 처리
            voiceRecognizer.AudioLevelChanged += VoiceRecognizer_AudioLevelChanged;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadMacros();
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.ShowInTaskbar = false; // 작업 표시줄에서 숨김
                trayIcon.BalloonTipTitle = "음성 매크로";
                trayIcon.BalloonTipText = "프로그램이 트레이로 최소화되었습니다. 더블 클릭하여 복원하세요.";
                trayIcon.ShowBalloonTip(3000);
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 프로그램이 완전히 종료될 때(폼이 닫히지만 트레이로 최소화하는 경우가 아닐 때)만 리소스 정리
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // 사용자에게 선택 옵션 제공
                DialogResult result = MessageBox.Show(
                    "음성 매크로 프로그램을 종료하시겠습니까?\n\n'예' - 프로그램 종료\n'아니오' - 시스템 트레이로 최소화",
                    "음성 매크로",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.No)
                {
                    // 시스템 트레이로 최소화
                    e.Cancel = true;  // 폼 닫기 취소
                    this.WindowState = FormWindowState.Minimized;
                    this.ShowInTaskbar = false;
                    return;
                }
            }

            try
            {
                // 프로그램 종료 시 리소스 정리
                // 리스닝 중이라면 먼저 중지
                if (isListening && voiceRecognizer != null)
                {
                    isListening = false;
                    try
                    {
                        voiceRecognizer.StopListening();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"음성인식 중지 중 오류: {ex.Message}");
                    }
                }
                
                // 그 다음 리소스 정리
                try 
                {
                    if (voiceRecognizer != null)
                    {
                        voiceRecognizer.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"음성인식 자원 해제 중 오류: {ex.Message}");
                }
                
                // 트레이 아이콘 정리
                if (trayIcon != null)
                {
                    try
                    {
                        trayIcon.Visible = false;
                        trayIcon.Dispose();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"트레이 아이콘 해제 중 오류: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // 종료 시 예외가 발생해도 사용자에게는 보여주지 않고 계속 종료 진행
                System.Diagnostics.Debug.WriteLine($"종료 중 오류 발생: {ex.Message}");
            }
            finally
            {
                // 항상 프로그램 종료 실행
                // Application.Exit();
                // Form.Close() 내부에서 이미 이 단계에 도달했으므로 추가로 Application.Exit()을 호출하지 않음
            }
        }

        private async void btnStartStop_Click(object sender, EventArgs e)
        {
            await ToggleListening();
        }

        private async Task ToggleListening()
        {
            try
            {
                if (isListening)
                {
                    // 리스닝 중지 전에 상태 먼저 변경
                    btnStartStop.Text = "시작";
                    isListening = false;
                    statusLabel.Text = "준비";
                    AddLogMessage("음성 인식이 중지되었습니다.", LogMessageType.Info);
                    
                    // 그 다음 음성 인식 서비스 중지
                    voiceRecognizer.StopListening();
                }
                else
                {
                    // 상태 먼저 변경
                    btnStartStop.Text = "중지";
                    isListening = true;
                    statusLabel.Text = "듣는 중...";
                    AddLogMessage("음성 인식이 시작되었습니다.", LogMessageType.Info);
                    
                    // 그 다음 음성 인식 시작
                    await voiceRecognizer.StartListening();
                }
            }
            catch (Exception ex)
            {
                // 예외 발생 시 상태 복원 및 오류 메시지 표시
                isListening = false;
                btnStartStop.Text = "시작";
                statusLabel.Text = "오류";
                AddLogMessage($"음성 인식 오류: {ex.Message}", LogMessageType.Error);
            }
        }

        private void btnAddMacro_Click(object sender, EventArgs e)
        {
            using (var addForm = new AddMacroForm(voiceRecognizer))
            {
                if (addForm.ShowDialog() == DialogResult.OK)
                {
                    // 매크로 키워드와 액션뿐만 아니라 액션 타입과 파라미터도 전달
                    macroService.AddMacro(
                        addForm.Keyword, 
                        addForm.KeyAction, 
                        addForm.SelectedActionType, 
                        addForm.SelectedActionParam
                    );
                    LoadMacros();
                    
                    // 매크로 추가 후 상태 메시지 표시
                    statusLabel.Text = $"매크로 '{addForm.Keyword}' 추가됨";
                    AddLogMessage($"새 매크로가 추가되었습니다. 키워드: {addForm.Keyword}, 액션 타입: {addForm.SelectedActionType}", LogMessageType.Info);
                }
            }
        }

        private void btnRemoveMacro_Click(object sender, EventArgs e)
        {
            if (lstMacros.SelectedItems.Count > 0)
            {
                string keyword = lstMacros.SelectedItems[0].Text;
                macroService.RemoveMacro(keyword);
                LoadMacros();
            }
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            // 음성 인식 일시 중지
            bool wasListening = isListening;
            if (wasListening)
            {
                ToggleListening();
            }

            // 설정 폼 표시
            using (var settingsForm = new SettingsForm(voiceRecognizer))
            {
                settingsForm.ShowDialog(this);
            }

            // 음성 인식 재개 (이전에 활성화되어 있었다면)
            if (wasListening)
            {
                ToggleListening();
            }
        }

        private void btnPresets_Click(object sender, EventArgs e)
        {
            // 프리셋 관리 대화상자 표시
            using (var presetForm = new PresetManagerForm(macroService))
            {
                if (presetForm.ShowDialog(this) == DialogResult.OK)
                {
                    // 프리셋이 로드되거나 가져와진 경우 목록 새로고침
                    LoadMacros();
                }
            }
        }

        private void LoadMacros()
        {
            lstMacros.Items.Clear();
            foreach (var macro in macroService.GetAllMacros())
            {
                var item = new ListViewItem(macro.Keyword);
                item.SubItems.Add(macro.KeyAction);
                
                // 액션 타입과 파라미터 정보 추가
                string actionTypeText = GetActionTypeDisplayText(macro.ActionType);
                item.SubItems.Add(actionTypeText);
                
                // 파라미터 정보가 0이 아닌 경우에만 표시
                if (macro.ActionParameters > 0)
                {
                    string paramText = GetActionParameterDisplayText(macro.ActionType, macro.ActionParameters);
                    item.SubItems.Add(paramText);
                }
                else
                {
                    item.SubItems.Add(string.Empty);
                }
                
                lstMacros.Items.Add(item);
            }
        }
        
        /// <summary>
        /// 액션 타입의 표시 텍스트를 반환합니다.
        /// </summary>
        /// <param name="actionType">액션 타입</param>
        /// <returns>표시할 텍스트</returns>
        private string GetActionTypeDisplayText(MacroActionType actionType)
        {
            switch (actionType)
            {
                case MacroActionType.Default:
                    return "기본";
                case MacroActionType.Toggle:
                    return "토글";
                case MacroActionType.Repeat:
                    return "반복";
                case MacroActionType.Hold:
                    return "홀드";
                case MacroActionType.Turbo:
                    return "터보";
                case MacroActionType.Combo:
                    return "콤보";
                default:
                    return "기본";
            }
        }
        
        /// <summary>
        /// 액션 파라미터의 표시 텍스트를 반환합니다.
        /// </summary>
        /// <param name="actionType">액션 타입</param>
        /// <param name="parameter">파라미터 값</param>
        /// <returns>표시할 텍스트</returns>
        private string GetActionParameterDisplayText(MacroActionType actionType, int parameter)
        {
            switch (actionType)
            {
                case MacroActionType.Repeat:
                    return $"{parameter}회";
                case MacroActionType.Hold:
                    return $"{parameter}ms ({parameter / 1000.0:F1}초)";
                case MacroActionType.Turbo:
                    return $"{parameter}ms 간격";
                case MacroActionType.Combo:
                    return $"{parameter}ms 간격";
                default:
                    return parameter.ToString();
            }
        }

        private void ChkDetailedLog_CheckedChanged(object sender, EventArgs e)
        {
            showDetailedLog = chkDetailedLog.Checked;
            AddLogMessage(showDetailedLog ? "상세 로그 표시가 활성화되었습니다." : "상세 로그 표시가 비활성화되었습니다.", LogMessageType.Info);
        }
        
        private void ChkPlayBeep_Click(object sender, EventArgs e)
        {
            playBeepSound = !playBeepSound;
            
            // 로그에 상태 변경 기록
            if (playBeepSound)
            {
                AddLogMessage("알림음 기능이 활성화되었습니다.", LogMessageType.Info);
                chkPlayBeep.Text = "알림음: 켜짐";
                chkPlayBeep.BackColor = Color.LightSkyBlue;
            }
            else
            {
                AddLogMessage("알림음 기능이 비활성화되었습니다.", LogMessageType.Info);
                chkPlayBeep.Text = "알림음: 꺼짐";
                chkPlayBeep.BackColor = Color.LightGray;
            }
        }

        private void btnCopyMacro_Click(object sender, EventArgs e)
        {
            if (lstMacros.SelectedItems.Count > 0)
            {
                string keyword = lstMacros.SelectedItems[0].Text;
                
                // 복사할 새 이름 입력받기
                string newKeyword = GetCopyKeywordFromUser(keyword);
                if (string.IsNullOrEmpty(newKeyword))
                {
                    return; // 사용자가 취소함
                }
                
                // 원본 매크로 찾기 - MacroService 사용하지 않고 직접 구현
                var allMacros = macroService.GetAllMacros();
                var sourceMacro = allMacros.FirstOrDefault(m => m.Keyword.Equals(keyword, StringComparison.OrdinalIgnoreCase));
                if (sourceMacro == null)
                {
                    MessageBox.Show($"복사할 매크로 '{keyword}'를 찾을 수 없습니다.", 
                        "매크로 복사 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // 새 키워드가 이미 존재하는지 확인
                if (allMacros.Any(m => m.Keyword.Equals(newKeyword, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show($"매크로 키워드 '{newKeyword}'가 이미 존재합니다.", 
                        "매크로 복사 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // 매크로 직접 추가
                macroService.AddMacro(
                    newKeyword, 
                    sourceMacro.KeyAction, 
                    sourceMacro.ActionType, 
                    sourceMacro.ActionParameters
                );
                
                LoadMacros(); // 목록 새로고침
                
                // 성공 메시지
                statusLabel.Text = $"매크로 '{keyword}'가 '{newKeyword}'로 복사되었습니다.";
                AddLogMessage($"매크로가 복사되었습니다. 키워드: {keyword} → {newKeyword}", LogMessageType.Info);
                
                // 새로 복사된 매크로 선택
                SelectMacroByKeyword(newKeyword);
            }
            else
            {
                MessageBox.Show("복사할 매크로를 선택해주세요.", "선택 필요", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        
        private void btnEditMacro_Click(object sender, EventArgs e)
        {
            if (lstMacros.SelectedItems.Count > 0)
            {
                string keyword = lstMacros.SelectedItems[0].Text;
                
                // 선택된 매크로 정보 가져오기 - 직접 구현
                var allMacros = macroService.GetAllMacros();
                var macro = allMacros.FirstOrDefault(m => m.Keyword.Equals(keyword, StringComparison.OrdinalIgnoreCase));
                if (macro == null)
                {
                    MessageBox.Show($"매크로 '{keyword}'를 찾을 수 없습니다.", 
                        "매크로 찾기 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // 매크로 편집 폼 열기
                using (var editForm = new AddMacroForm(voiceRecognizer, 
                    macro.Keyword, macro.KeyAction, macro.ActionType, macro.ActionParameters))
                {
                    if (editForm.ShowDialog() == DialogResult.OK)
                    {
                        // 기존 매크로 삭제
                        macroService.RemoveMacro(editForm.OriginalKeyword);
                        
                        // 새 매크로 추가
                        macroService.AddMacro(
                            editForm.Keyword, 
                            editForm.KeyAction, 
                            editForm.SelectedActionType, 
                            editForm.SelectedActionParam
                        );
                        
                        LoadMacros(); // 목록 새로고침
                        
                        // 성공 메시지
                        statusLabel.Text = $"매크로 '{editForm.OriginalKeyword}'가 업데이트되었습니다.";
                        AddLogMessage($"매크로가 수정되었습니다. 키워드: {editForm.OriginalKeyword} → {editForm.Keyword}", 
                            LogMessageType.Info);
                        
                        // 편집된 매크로 선택
                        SelectMacroByKeyword(editForm.Keyword);
                    }
                }
            }
            else
            {
                MessageBox.Show("수정할 매크로를 선택해주세요.", "선택 필요", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        
        /// <summary>
        /// 매크로 복사 시 새로운 키워드를 입력받는 대화상자를 표시합니다.
        /// </summary>
        /// <param name="originalKeyword">원본 매크로 키워드</param>
        /// <returns>새 키워드 또는 취소 시 null</returns>
        private string GetCopyKeywordFromUser(string originalKeyword)
        {
            using (var inputForm = new Form())
            {
                inputForm.Text = "매크로 복사";
                inputForm.Size = new Size(400, 150);
                inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputForm.StartPosition = FormStartPosition.CenterParent;
                inputForm.MaximizeBox = false;
                inputForm.MinimizeBox = false;
                
                // 안내 레이블
                var label = new Label
                {
                    Text = "새 매크로 이름을 입력하세요:",
                    Location = new Point(10, 20),
                    Size = new Size(380, 20)
                };
                
                // 텍스트 입력 상자
                var textBox = new TextBox
                {
                    Text = $"{originalKeyword}_복사본",
                    Location = new Point(10, 50),
                    Size = new Size(360, 20)
                };
                
                // OK 버튼
                var okButton = new Button
                {
                    Text = "확인",
                    DialogResult = DialogResult.OK,
                    Location = new Point(200, 80),
                    Size = new Size(80, 30)
                };
                
                // 취소 버튼
                var cancelButton = new Button
                {
                    Text = "취소",
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(290, 80),
                    Size = new Size(80, 30)
                };
                
                // 컨트롤 추가 및 기본 버튼 설정
                inputForm.Controls.Add(label);
                inputForm.Controls.Add(textBox);
                inputForm.Controls.Add(okButton);
                inputForm.Controls.Add(cancelButton);
                inputForm.AcceptButton = okButton;
                inputForm.CancelButton = cancelButton;
                
                // 대화상자 표시 및 결과 반환
                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    return textBox.Text.Trim();
                }
                
                return null; // 취소됨
            }
        }
        
        /// <summary>
        /// 지정된 키워드의 매크로를 리스트뷰에서 선택합니다.
        /// </summary>
        /// <param name="keyword">선택할 매크로 키워드</param>
        private void SelectMacroByKeyword(string keyword)
        {
            foreach (ListViewItem item in lstMacros.Items)
            {
                if (item.Text == keyword)
                {
                    item.Selected = true;
                    item.EnsureVisible();
                    break;
                }
            }
        }

        /// <summary>
        /// 마이크 레벨이 변경될 때 호출되는 이벤트 핸들러
        /// </summary>
        private void VoiceRecognizer_AudioLevelChanged(object sender, float level)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => VoiceRecognizer_AudioLevelChanged(sender, level)));
                return;
            }
            
            // 현재 마이크 레벨을 프로그레스바에 표시
            int levelValue = (int)(level * 100);
            pbarMicLevel.Value = Math.Min(levelValue, 100);
            
            // 레벨에 따라 색상 변경
            if (levelValue < 30)
            {
                pbarMicLevel.BarColor = Color.Green;
            }
            else if (levelValue < 70)
            {
                pbarMicLevel.BarColor = Color.Yellow;
            }
            else
            {
                pbarMicLevel.BarColor = Color.Red;
            }
        }

        /// <summary>
        /// 마이크 볼륨 슬라이더의 값이 변경될 때 호출되는 이벤트 핸들러
        /// </summary>
        private void TbarMicVolume_ValueChanged(object sender, EventArgs e)
        {
            float volume = tbarMicVolume.Value / 100.0f;
            
            try
            {
                // 오디오 레코딩 서비스에 볼륨 설정
                if (voiceRecognizer != null)
                {
                    // VoiceRecognitionService를 통해 AudioRecordingService에 접근
                    var audioRecordingService = voiceRecognizer.GetAudioRecordingService();
                    if (audioRecordingService != null)
                    {
                        audioRecordingService.SetMicrophoneVolume(volume);
                    }
                }
            }
            catch (Exception ex)
            {
                AddLogMessage($"마이크 볼륨 설정 오류: {ex.Message}", LogMessageType.Error);
            }
        }
    }
} 