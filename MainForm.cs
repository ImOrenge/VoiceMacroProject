using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VoiceMacro.Services;

namespace VoiceMacro
{
    public partial class MainForm : Form
    {
        private NotifyIcon trayIcon;
        private VoiceRecognitionService voiceRecognizer;
        private MacroService macroService;
        private bool isListening = false;
        private RichTextBox rtbLog;
        private Button btnClearLog;
        private CheckBox chkAutoScroll;

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
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // MainForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 550); // 폼 높이 증가
            this.Name = "MainForm";
            this.Text = "음성 매크로";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Resize += new System.EventHandler(this.MainForm_Resize);
            this.ResumeLayout(false);

            // 버튼, 목록 및 기타 컨트롤 추가
            this.btnStartStop = new Button();
            this.btnStartStop.Location = new Point(20, 20);
            this.btnStartStop.Size = new Size(100, 30);
            this.btnStartStop.Text = "시작";
            this.btnStartStop.Click += new EventHandler(this.btnStartStop_Click);
            this.Controls.Add(this.btnStartStop);

            this.btnAddMacro = new Button();
            this.btnAddMacro.Location = new Point(20, 60);
            this.btnAddMacro.Size = new Size(100, 30);
            this.btnAddMacro.Text = "매크로 추가";
            this.btnAddMacro.Click += new EventHandler(this.btnAddMacro_Click);
            this.Controls.Add(this.btnAddMacro);

            this.btnRemoveMacro = new Button();
            this.btnRemoveMacro.Location = new Point(130, 60);
            this.btnRemoveMacro.Size = new Size(100, 30);
            this.btnRemoveMacro.Text = "매크로 삭제";
            this.btnRemoveMacro.Click += new EventHandler(this.btnRemoveMacro_Click);
            this.Controls.Add(this.btnRemoveMacro);

            this.btnSettings = new Button();
            this.btnSettings.Location = new Point(240, 60);
            this.btnSettings.Size = new Size(100, 30);
            this.btnSettings.Text = "설정";
            this.btnSettings.Click += new EventHandler(this.btnSettings_Click);
            this.Controls.Add(this.btnSettings);

            this.btnPresets = new Button();
            this.btnPresets.Location = new Point(350, 60);
            this.btnPresets.Size = new Size(100, 30);
            this.btnPresets.Text = "프리셋";
            this.btnPresets.Click += new EventHandler(this.btnPresets_Click);
            this.Controls.Add(this.btnPresets);

            this.lstMacros = new ListView();
            this.lstMacros.Location = new Point(20, 100);
            this.lstMacros.Size = new Size(560, 250);
            this.lstMacros.View = View.Details;
            this.lstMacros.FullRowSelect = true;
            this.lstMacros.Columns.Add("명령어", 120);
            this.lstMacros.Columns.Add("동작", 400);
            this.Controls.Add(this.lstMacros);

            // 로그 영역 레이블
            Label lblLogArea = new Label();
            lblLogArea.Text = "실행 로그";
            lblLogArea.Location = new Point(20, 365);
            lblLogArea.Size = new Size(100, 20);
            lblLogArea.Font = new Font(lblLogArea.Font, FontStyle.Bold);
            this.Controls.Add(lblLogArea);

            // 로그 컨트롤 버튼들
            this.btnClearLog = new Button();
            this.btnClearLog.Text = "로그 지우기";
            this.btnClearLog.Location = new Point(480, 360);
            this.btnClearLog.Size = new Size(100, 25);
            this.btnClearLog.Click += BtnClearLog_Click;
            this.Controls.Add(this.btnClearLog);

            this.chkAutoScroll = new CheckBox();
            this.chkAutoScroll.Text = "자동 스크롤";
            this.chkAutoScroll.Checked = true;
            this.chkAutoScroll.Location = new Point(370, 363);
            this.chkAutoScroll.Size = new Size(100, 20);
            this.Controls.Add(this.chkAutoScroll);

            // 로그 메시지 표시 RichTextBox
            this.rtbLog = new RichTextBox();
            this.rtbLog.Location = new Point(20, 390);
            this.rtbLog.Size = new Size(560, 120);
            this.rtbLog.ReadOnly = true;
            this.rtbLog.BackColor = Color.White;
            this.rtbLog.Font = new Font("Consolas", 9F);
            this.rtbLog.ScrollBars = RichTextBoxScrollBars.Vertical;
            this.Controls.Add(this.rtbLog);

            this.statusStrip = new StatusStrip();
            this.statusLabel = new ToolStripStatusLabel();
            this.statusLabel.Text = "준비";
            this.statusStrip.Items.Add(this.statusLabel);
            this.Controls.Add(this.statusStrip);
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
            AddLogMessage($"상태: {status}", LogMessageType.Info);
        }

        private void VoiceRecognizer_SpeechRecognized(object? sender, string text)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => VoiceRecognizer_SpeechRecognized(sender, text)));
                return;
            }

            AddLogMessage($"인식됨: {text}", LogMessageType.Recognition);
        }

        private void MacroService_StatusChanged(object? sender, string status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => MacroService_StatusChanged(sender, status)));
                return;
            }

            AddLogMessage($"매크로: {status}", LogMessageType.Info);
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
            }
            else
            {
                AddLogMessage($"실패: '{e.Keyword}' → {e.KeyAction} ({e.ErrorMessage})", LogMessageType.Error);
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

        private Button btnStartStop;
        private Button btnAddMacro;
        private Button btnRemoveMacro;
        private Button btnSettings;
        private Button btnPresets;
        private ListView lstMacros;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;

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
                AddLogMessage("프로그램이 트레이에서 복원되었습니다.", LogMessageType.Info);
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
                // 리소스 정리 후 종료
                voiceRecognizer?.StopListening();
                voiceRecognizer?.Dispose();
                trayIcon.Visible = false;
                Application.Exit(); 
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
                    AddLogMessage("프로그램이 트레이에서 복원되었습니다.", LogMessageType.Info);
                }
            };
        }

        private void InitializeServices()
        {
            macroService = new MacroService();
            voiceRecognizer = new VoiceRecognitionService(macroService);
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
                AddLogMessage("프로그램이 시스템 트레이로 최소화되었습니다.", LogMessageType.Info);
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
                    AddLogMessage("프로그램이 시스템 트레이로 최소화되었습니다.", LogMessageType.Info);
                    return;
                }
            }

            // 프로그램 종료 시 리소스 정리
            voiceRecognizer?.StopListening();
            voiceRecognizer?.Dispose();
            trayIcon?.Dispose();

            // 설정 저장 등 필요한 정리 작업 수행
            Application.Exit();
        }

        private async void btnStartStop_Click(object sender, EventArgs e)
        {
            await ToggleListening();
        }

        private async Task ToggleListening()
        {
            if (isListening)
            {
                voiceRecognizer.StopListening();
                btnStartStop.Text = "시작";
                isListening = false;
                statusLabel.Text = "준비";
                AddLogMessage("음성 인식이 중지되었습니다.", LogMessageType.Info);
            }
            else
            {
                btnStartStop.Text = "중지";
                isListening = true;
                statusLabel.Text = "듣는 중...";
                AddLogMessage("음성 인식이 시작되었습니다.", LogMessageType.Info);
                await voiceRecognizer.StartListening();
            }
        }

        private void btnAddMacro_Click(object sender, EventArgs e)
        {
            using (var addForm = new AddMacroForm(voiceRecognizer))
            {
                if (addForm.ShowDialog() == DialogResult.OK)
                {
                    macroService.AddMacro(addForm.Keyword, addForm.KeyAction);
                    LoadMacros();
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
                lstMacros.Items.Add(item);
            }
        }
    }
} 