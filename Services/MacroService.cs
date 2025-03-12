using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;

namespace VoiceMacro.Services
{
    /// <summary>
    /// 매크로 액션 유형을 정의하는 열거형입니다.
    /// </summary>
    public enum MacroActionType
    {
        /// <summary>
        /// 기본 키 액션: 한 번 입력
        /// </summary>
        Default = 0,
        
        /// <summary>
        /// 토글 키 액션: 키를 누르고 떼는 동작 전환
        /// </summary>
        Toggle = 1,
        
        /// <summary>
        /// n차 반복 액션: 지정된 횟수만큼 반복
        /// </summary>
        Repeat = 2,
        
        /// <summary>
        /// n초 동안 입력 유지: 지정된 시간 동안 키를 누르고 있음
        /// </summary>
        Hold = 3,
        
        /// <summary>
        /// 터보(빠른 키 연타): 빠른 속도로 키를 반복해서 누름
        /// </summary>
        Turbo = 4,
        
        /// <summary>
        /// 콤보(키 순차 입력): 여러 키를 순차적으로 입력
        /// </summary>
        Combo = 5
    }

    /// <summary>
    /// 매크로 실행 이벤트 인자 클래스입니다.
    /// 매크로 실행 결과와 관련된 정보를 전달합니다.
    /// </summary>
    public class MacroExecutionEventArgs : EventArgs
    {
        /// <summary>
        /// 실행된 매크로의 키워드입니다.
        /// </summary>
        public string Keyword { get; }

        /// <summary>
        /// 실행된 키보드 동작입니다.
        /// </summary>
        public string KeyAction { get; }

        /// <summary>
        /// 매크로 실행 성공 여부입니다.
        /// </summary>
        public bool Success { get; }

        /// <summary>
        /// 실패시 오류 메시지입니다. 성공 시에는 null입니다.
        /// </summary>
        public string? ErrorMessage { get; }

        /// <summary>
        /// 매크로 실행 이벤트 인자 생성자입니다.
        /// </summary>
        /// <param name="keyword">실행된 매크로 키워드</param>
        /// <param name="keyAction">실행된 키보드 동작</param>
        /// <param name="success">실행 성공 여부</param>
        /// <param name="errorMessage">오류 메시지 (실패 시)</param>
        public MacroExecutionEventArgs(string keyword, string keyAction, bool success, string? errorMessage = null)
        {
            Keyword = keyword;
            KeyAction = keyAction;
            Success = success;
            ErrorMessage = errorMessage;
        }
    }

    /// <summary>
    /// 매크로 서비스 클래스입니다.
    /// 매크로 추가, 삭제, 저장, 로드 및 실행을 담당합니다.
    /// </summary>
    public class MacroService
    {
        /// <summary>
        /// 매크로 실행 완료 시 발생하는 이벤트입니다.
        /// </summary>
        public event EventHandler<MacroExecutionEventArgs>? MacroExecuted;
        
        /// <summary>
        /// 매크로 서비스 상태 변경 시 발생하는 이벤트입니다.
        /// </summary>
        public event EventHandler<string>? StatusChanged;

        /// <summary>
        /// 매크로 목록을 저장하는 파일 경로입니다.
        /// </summary>
        private readonly string macroFilePath;

        /// <summary>
        /// 프리셋 파일을 저장하는 디렉토리 경로입니다.
        /// </summary>
        private readonly string presetDirectory;

        /// <summary>
        /// 키보드 입력 시뮬레이션을 위한 객체입니다.
        /// </summary>
        private readonly InputSimulator inputSimulator;

        /// <summary>
        /// 매크로 목록입니다.
        /// </summary>
        private List<Macro> macros;

        /// <summary>
        /// 토글 키 액션을 위한 키 상태 저장 변수입니다.
        /// </summary>
        private readonly Dictionary<string, bool> keyStates = new Dictionary<string, bool>();

        /// <summary>
        /// 매크로 서비스 생성자입니다.
        /// </summary>
        public MacroService()
        {
            // 데이터 디렉토리 설정
            string appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VoiceMacro");
            Directory.CreateDirectory(appDataDir);
            
            // 매크로 파일 경로 설정
            macroFilePath = Path.Combine(appDataDir, "macros.json");
            
            // 프리셋 디렉토리 설정
            presetDirectory = Path.Combine(appDataDir, "Presets");
            Directory.CreateDirectory(presetDirectory);
            
            // 입력 시뮬레이터 초기화
            inputSimulator = new InputSimulator();
            
            // 매크로 로드
            LoadMacros();
        }

        /// <summary>
        /// 저장된 매크로 목록을 가져옵니다.
        /// </summary>
        public List<Macro> GetMacros()
        {
            return macros;
        }

        /// <summary>
        /// 매크로를 추가합니다.
        /// </summary>
        /// <param name="keyword">음성 키워드</param>
        /// <param name="keyAction">실행할 키보드 동작</param>
        public void AddMacro(string keyword, string keyAction)
        {
            if (string.IsNullOrWhiteSpace(keyword) || string.IsNullOrWhiteSpace(keyAction))
            {
                OnStatusChanged("키워드와 키 동작은 비워둘 수 없습니다.");
                return;
            }

            // 기존 매크로 확인
            var existingMacro = macros.FirstOrDefault(m => m.Keyword.Equals(keyword, StringComparison.OrdinalIgnoreCase));
            
            if (existingMacro != null)
            {
                // 기존 매크로 업데이트
                existingMacro.KeyAction = keyAction;
                OnStatusChanged($"매크로 '{keyword}' 업데이트됨");
            }
            else
            {
                // 새 매크로 추가
                macros.Add(new Macro { Keyword = keyword, KeyAction = keyAction });
                OnStatusChanged($"매크로 '{keyword}' 추가됨");
            }
            
            // 변경사항 저장
            SaveMacros();
        }

        /// <summary>
        /// 매크로를 추가합니다. (액션 타입과 파라미터 포함)
        /// </summary>
        /// <param name="keyword">음성 키워드</param>
        /// <param name="keyAction">실행할 키보드 동작</param>
        /// <param name="actionType">매크로 액션 유형</param>
        /// <param name="actionParam">매크로 액션 파라미터</param>
        public void AddMacro(string keyword, string keyAction, MacroActionType actionType, int actionParam)
        {
            if (string.IsNullOrWhiteSpace(keyword) || string.IsNullOrWhiteSpace(keyAction))
            {
                OnStatusChanged("키워드와 키 동작은 비워둘 수 없습니다.");
                return;
            }

            // 기존 매크로 확인
            var existingMacro = macros.FirstOrDefault(m => m.Keyword.Equals(keyword, StringComparison.OrdinalIgnoreCase));
            
            if (existingMacro != null)
            {
                // 기존 매크로 업데이트
                existingMacro.KeyAction = keyAction;
                existingMacro.ActionType = actionType;
                existingMacro.ActionParameters = actionParam;
                OnStatusChanged($"매크로 '{keyword}' 업데이트됨 (액션 타입: {actionType})");
            }
            else
            {
                // 새 매크로 추가
                macros.Add(new Macro 
                { 
                    Keyword = keyword, 
                    KeyAction = keyAction,
                    ActionType = actionType,
                    ActionParameters = actionParam
                });
                OnStatusChanged($"매크로 '{keyword}' 추가됨 (액션 타입: {actionType})");
            }
            
            // 변경사항 저장
            SaveMacros();
        }

        /// <summary>
        /// 매크로를 삭제합니다.
        /// </summary>
        /// <param name="keyword">삭제할 매크로의 키워드</param>
        public void RemoveMacro(string keyword)
        {
            macros.RemoveAll(m => m.Keyword.Equals(keyword, StringComparison.OrdinalIgnoreCase));
            SaveMacros();
            OnStatusChanged($"매크로 삭제됨: {keyword}");
        }

        /// <summary>
        /// 기존 매크로를 복사하여 새 매크로를 생성합니다.
        /// </summary>
        /// <param name="sourceKeyword">원본 매크로 키워드</param>
        /// <param name="newKeyword">새 매크로 키워드</param>
        /// <returns>복사 성공 여부</returns>
        public bool CopyMacro(string sourceKeyword, string newKeyword)
        {
            // 원본 매크로 찾기
            var sourceMacro = macros.FirstOrDefault(m => m.Keyword.Equals(sourceKeyword, StringComparison.OrdinalIgnoreCase));
            if (sourceMacro == null)
            {
                OnStatusChanged($"복사할 매크로 '{sourceKeyword}'를 찾을 수 없습니다.");
                return false;
            }
            
            // 새 키워드가 이미 존재하는지 확인
            if (macros.Any(m => m.Keyword.Equals(newKeyword, StringComparison.OrdinalIgnoreCase)))
            {
                OnStatusChanged($"매크로 키워드 '{newKeyword}'가 이미 존재합니다.");
                return false;
            }
            
            // 새 매크로 생성 및 추가
            var newMacro = new Macro
            {
                Keyword = newKeyword,
                KeyAction = sourceMacro.KeyAction,
                ActionType = sourceMacro.ActionType,
                ActionParameters = sourceMacro.ActionParameters
            };
            
            macros.Add(newMacro);
            SaveMacros();
            OnStatusChanged($"매크로 '{sourceKeyword}'가 '{newKeyword}'로 복사되었습니다.");
            return true;
        }
        
        /// <summary>
        /// 기존 매크로를 수정합니다.
        /// </summary>
        /// <param name="originalKeyword">원본 매크로 키워드</param>
        /// <param name="newKeyword">새 매크로 키워드</param>
        /// <param name="keyAction">새 키 동작</param>
        /// <param name="actionType">새 액션 타입</param>
        /// <param name="actionParam">새 액션 파라미터</param>
        /// <returns>수정 성공 여부</returns>
        public bool UpdateMacro(string originalKeyword, string newKeyword, string keyAction, 
                                MacroActionType actionType, int actionParam)
        {
            // 기존 매크로 찾기
            var macro = macros.FirstOrDefault(m => m.Keyword.Equals(originalKeyword, StringComparison.OrdinalIgnoreCase));
            if (macro == null)
            {
                OnStatusChanged($"수정할 매크로 '{originalKeyword}'를 찾을 수 없습니다.");
                return false;
            }
            
            // 키워드가 변경되었고, 새 키워드가 이미 존재하는지 확인
            if (!originalKeyword.Equals(newKeyword, StringComparison.OrdinalIgnoreCase) && 
                macros.Any(m => m.Keyword.Equals(newKeyword, StringComparison.OrdinalIgnoreCase)))
            {
                OnStatusChanged($"매크로 키워드 '{newKeyword}'가 이미 존재합니다.");
                return false;
            }
            
            // 매크로 업데이트
            macro.Keyword = newKeyword;
            macro.KeyAction = keyAction;
            macro.ActionType = actionType;
            macro.ActionParameters = actionParam;
            
            SaveMacros();
            OnStatusChanged($"매크로 '{originalKeyword}'가 업데이트되었습니다.");
            return true;
        }
        
        /// <summary>
        /// 지정된 키워드에 해당하는 매크로를 가져옵니다.
        /// </summary>
        /// <param name="keyword">매크로 키워드</param>
        /// <returns>매크로 객체 또는 null(찾지 못한 경우)</returns>
        public Macro? GetMacro(string keyword)
        {
            return macros.FirstOrDefault(m => m.Keyword.Equals(keyword, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 프리셋으로 현재 매크로 목록을 저장합니다.
        /// </summary>
        /// <param name="presetName">프리셋 이름</param>
        /// <returns>저장 성공 여부</returns>
        public bool SavePreset(string presetName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(presetName))
                {
                    OnStatusChanged("프리셋 이름이 유효하지 않습니다.");
                    return false;
                }
                
                // 유효한 파일 이름으로 변환
                string validFileName = string.Join("_", presetName.Split(Path.GetInvalidFileNameChars()));
                string presetFilePath = Path.Combine(presetDirectory, $"{validFileName}.json");
                
                // 매크로 목록을 JSON으로 직렬화하여 저장
                string json = JsonSerializer.Serialize(macros, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(presetFilePath, json);
                
                OnStatusChanged($"프리셋 '{presetName}' 저장 완료");
                return true;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"프리셋 저장 오류: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 프리셋을 불러와 현재 매크로 목록을 교체합니다.
        /// </summary>
        /// <param name="presetName">프리셋 이름</param>
        /// <returns>로드 성공 여부</returns>
        public bool LoadPreset(string presetName)
        {
            try
            {
                string presetFilePath = Path.Combine(presetDirectory, $"{presetName}.json");
                
                if (!File.Exists(presetFilePath))
                {
                    OnStatusChanged("프리셋 파일이 존재하지 않습니다.");
                    return false;
                }
                
                // 프리셋 파일에서 매크로 목록 로드
                string json = File.ReadAllText(presetFilePath);
                List<Macro> loadedMacros = JsonSerializer.Deserialize<List<Macro>>(json);
                
                if (loadedMacros == null || loadedMacros.Count == 0)
                {
                    OnStatusChanged("프리셋에 매크로가 없습니다.");
                    return false;
                }
                
                // 현재 매크로 목록 교체 및 저장
                macros = loadedMacros;
                SaveMacros();
                
                OnStatusChanged($"프리셋 '{presetName}' 불러오기 완료 ({macros.Count}개 매크로)");
                return true;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"프리셋 불러오기 오류: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 음성 명령 처리 메서드입니다.
        /// </summary>
        /// <param name="command">처리할 음성 명령어</param>
        /// <returns>매크로 실행 성공 여부</returns>
        public bool ProcessVoiceCommand(string command)
        {
            if (string.IsNullOrEmpty(command))
            {
                // 빈 명령어는 무시만 하고 로그 메시지 발생하지 않음
                return false;
            }

            // 기존 ExecuteMacroIfMatched 메서드의 로직을 사용
            return ExecuteMacroIfMatched(command);
        }

        /// <summary>
        /// 인식된 음성에서 매크로를 찾아 실행합니다.
        /// </summary>
        /// <param name="recognizedText">인식된 텍스트</param>
        /// <returns>매크로 실행 성공 여부</returns>
        public bool ExecuteMacroIfMatched(string recognizedText)
        {
            string normalizedText = recognizedText.ToLower().Trim();
            
            foreach (var macro in macros)
            {
                string normalizedKeyword = macro.Keyword.ToLower().Trim();
                
                // 키워드가 인식된 텍스트에 포함되어 있는지 확인
                if (normalizedText.Contains(normalizedKeyword))
                {
                    try
                    {
                        // 키 동작 실행
                        ExecuteKeyAction(macro);
                        
                        // 실행 완료 이벤트 발생
                        OnMacroExecuted(macro.Keyword, macro.KeyAction, true);
                        OnStatusChanged($"매크로 '{macro.Keyword}' 실행 완료");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        OnMacroExecuted(macro.Keyword, macro.KeyAction, false, ex.Message);
                        OnStatusChanged($"매크로 '{macro.Keyword}' 실행 실패: {ex.Message}");
                    }
                }
            }
            
            return false;
        }

        /// <summary>
        /// 키보드 동작을 실행합니다.
        /// </summary>
        /// <param name="macro">실행할 매크로 객체</param>
        private void ExecuteKeyAction(Macro macro)
        {
            try
            {
                string keyAction = macro.KeyAction;
                
                // 액션 유형에 따라 다른 방식으로 키 작동 실행
                switch (macro.ActionType)
                {
                    case MacroActionType.Default:
                        // 기본 키 액션: 한 번 입력
                        ExecuteDefaultKeyAction(keyAction);
                        break;
                        
                    case MacroActionType.Toggle:
                        // 토글 키 액션: 키를 누르고 떼는 동작 전환
                        ExecuteToggleKeyAction(keyAction);
                        break;
                        
                    case MacroActionType.Repeat:
                        // n차 반복 액션: 지정된 횟수만큼 반복
                        int repeatCount = macro.ActionParameters > 0 ? macro.ActionParameters : 3; // 기본값 3회
                        ExecuteRepeatKeyAction(keyAction, repeatCount);
                        break;
                        
                    case MacroActionType.Hold:
                        // n초 동안 입력 유지: 지정된 시간 동안 키를 누르고 있음
                        int holdTime = macro.ActionParameters > 0 ? macro.ActionParameters : 1000; // 기본값 1초
                        ExecuteHoldKeyAction(keyAction, holdTime);
                        break;
                        
                    case MacroActionType.Turbo:
                        // 터보(빠른 키 연타): 빠른 속도로 키를 반복해서 누름
                        int turboSpeed = macro.ActionParameters > 0 ? macro.ActionParameters : 50; // 기본값 50ms 간격
                        int turboCount = 10; // 기본 10회 연타
                        ExecuteTurboKeyAction(keyAction, turboSpeed, turboCount);
                        break;
                        
                    case MacroActionType.Combo:
                        // 콤보(키 순차 입력): 여러 키를 순차적으로 입력
                        int comboDelay = macro.ActionParameters > 0 ? macro.ActionParameters : 100; // 기본값 100ms 간격
                        ExecuteComboKeyAction(keyAction, comboDelay);
                        break;
                        
                    default:
                        ExecuteDefaultKeyAction(keyAction);
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"키 동작 실행 중 오류 발생: {ex.Message}");
                throw; // 예외를 상위로 전파하여 ExecuteMacroIfMatched에서 처리하도록 함
            }
        }
        
        /// <summary>
        /// 기본 키 액션을 실행합니다.
        /// </summary>
        /// <param name="keyAction">실행할 키보드 동작 문자열</param>
        private void ExecuteDefaultKeyAction(string keyAction)
        {
            // Windows.Forms.SendKeys 형식의 키 조합
            if (keyAction.StartsWith("{") || keyAction.Contains("+") || keyAction.Contains("^") || keyAction.Contains("%"))
            {
                SendKeys.SendWait(keyAction);
            }
            // 특수 키워드 처리
            else if (keyAction.Equals("ESC", StringComparison.OrdinalIgnoreCase) || 
                     keyAction.Equals("ESCAPE", StringComparison.OrdinalIgnoreCase))
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.ESCAPE);
            }
            // 기타 InputSimulator로 처리
            else
            {
                foreach (char c in keyAction)
                {
                    inputSimulator.Keyboard.TextEntry(c);
                }
            }
        }
        
        /// <summary>
        /// 토글 키 액션을 실행합니다.
        /// </summary>
        /// <param name="keyAction">실행할 키보드 동작 문자열</param>
        private void ExecuteToggleKeyAction(string keyAction)
        {
            // 입력된 키를 누르고 있는 상태인지 확인하는 정적 변수 (간단한 구현을 위해 딕셔너리 사용)
            // 실제 애플리케이션에서는 이 상태를 더 안정적으로 관리해야 합니다
            if (!keyStates.ContainsKey(keyAction))
            {
                keyStates[keyAction] = false;
            }
            
            // 현재 상태의 반대로 전환
            bool isKeyDown = keyStates[keyAction];
            
            if (isKeyDown)
            {
                // 키가 눌려 있으면 키 누름 해제
                ReleaseKey(keyAction);
                keyStates[keyAction] = false;
            }
            else
            {
                // 키가 눌려 있지 않으면 키 누름
                PressKey(keyAction);
                keyStates[keyAction] = true;
            }
        }
        
        /// <summary>
        /// n차 반복 액션을 실행합니다.
        /// </summary>
        /// <param name="keyAction">실행할 키보드 동작 문자열</param>
        /// <param name="repeatCount">반복 횟수</param>
        private void ExecuteRepeatKeyAction(string keyAction, int repeatCount)
        {
            for (int i = 0; i < repeatCount; i++)
            {
                ExecuteDefaultKeyAction(keyAction);
                Thread.Sleep(50); // 반복 사이의 짧은 지연
            }
        }
        
        /// <summary>
        /// n초 동안 입력 유지 액션을 실행합니다.
        /// </summary>
        /// <param name="keyAction">실행할 키보드 동작 문자열</param>
        /// <param name="holdTimeMs">유지 시간(밀리초)</param>
        private void ExecuteHoldKeyAction(string keyAction, int holdTimeMs)
        {
            // 키 누르기
            PressKey(keyAction);
            
            // 지정된 시간 동안 대기
            Thread.Sleep(holdTimeMs);
            
            // 키 떼기
            ReleaseKey(keyAction);
        }
        
        /// <summary>
        /// 터보(빠른 키 연타) 액션을 실행합니다.
        /// </summary>
        /// <param name="keyAction">실행할 키보드 동작 문자열</param>
        /// <param name="turboSpeedMs">연타 간격(밀리초)</param>
        /// <param name="turboCount">연타 횟수</param>
        private void ExecuteTurboKeyAction(string keyAction, int turboSpeedMs, int turboCount)
        {
            for (int i = 0; i < turboCount; i++)
            {
                ExecuteDefaultKeyAction(keyAction);
                Thread.Sleep(turboSpeedMs);
            }
        }
        
        /// <summary>
        /// 콤보(키 순차 입력) 액션을 실행합니다.
        /// </summary>
        /// <param name="keyAction">실행할 키보드 동작 문자열(쉼표로 구분된 여러 키)</param>
        /// <param name="comboDelayMs">키 사이 간격(밀리초)</param>
        private void ExecuteComboKeyAction(string keyAction, int comboDelayMs)
        {
            // 쉼표로 구분된 키 목록 파싱
            string[] keys = keyAction.Split(',');
            
            foreach (string key in keys)
            {
                ExecuteDefaultKeyAction(key.Trim());
                Thread.Sleep(comboDelayMs);
            }
        }
        
        /// <summary>
        /// 키를 누르는 동작을 실행합니다.
        /// </summary>
        /// <param name="keyAction">실행할 키보드 동작 문자열</param>
        private void PressKey(string keyAction)
        {
            // SendKeys 형식의 키 조합은 InputSimulator로 변환하여 처리해야 함
            if (keyAction.StartsWith("{") || keyAction.Contains("+") || keyAction.Contains("^") || keyAction.Contains("%"))
            {
                // SendKeys는 누르고 있기 기능을 지원하지 않음
                // 여기서는 간단하게 VirtualKeyCode로 변환 가능한 일부 키만 처리
                if (keyAction.Equals("{ENTER}", StringComparison.OrdinalIgnoreCase))
                {
                    inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
                }
                else if (keyAction.Equals("{SPACE}", StringComparison.OrdinalIgnoreCase))
                {
                    inputSimulator.Keyboard.KeyDown(VirtualKeyCode.SPACE);
                }
                else if (keyAction.Equals("^c", StringComparison.OrdinalIgnoreCase)) // Ctrl+C
                {
                    inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_C);
                }
                else
                {
                    // 지원하지 않는 키 조합은 기본 동작으로 대체
                    ExecuteDefaultKeyAction(keyAction);
                }
            }
            else if (keyAction.Length == 1)
            {
                // 단일 문자 키
                char key = keyAction[0];
                VirtualKeyCode vkCode = GetVirtualKeyCodeFromChar(key);
                inputSimulator.Keyboard.KeyDown(vkCode);
            }
            else if (keyAction.Equals("ESC", StringComparison.OrdinalIgnoreCase) || keyAction.Equals("ESCAPE", StringComparison.OrdinalIgnoreCase))
            {
                inputSimulator.Keyboard.KeyDown(VirtualKeyCode.ESCAPE);
            }
            else
            {
                // 지원하지 않는 키는 기본 동작으로 대체
                ExecuteDefaultKeyAction(keyAction);
            }
        }
        
        /// <summary>
        /// 키를 떼는 동작을 실행합니다.
        /// </summary>
        /// <param name="keyAction">실행할 키보드 동작 문자열</param>
        private void ReleaseKey(string keyAction)
        {
            if (keyAction.StartsWith("{") || keyAction.Contains("+") || keyAction.Contains("^") || keyAction.Contains("%"))
            {
                // SendKeys 형식의 키 조합 처리
                if (keyAction.Equals("{ENTER}", StringComparison.OrdinalIgnoreCase))
                {
                    inputSimulator.Keyboard.KeyUp(VirtualKeyCode.RETURN);
                }
                else if (keyAction.Equals("{SPACE}", StringComparison.OrdinalIgnoreCase))
                {
                    inputSimulator.Keyboard.KeyUp(VirtualKeyCode.SPACE);
                }
                // 다른 조합키는 여기서 추가 처리
            }
            else if (keyAction.Length == 1)
            {
                // 단일 문자 키
                char key = keyAction[0];
                VirtualKeyCode vkCode = GetVirtualKeyCodeFromChar(key);
                inputSimulator.Keyboard.KeyUp(vkCode);
            }
            else if (keyAction.Equals("ESC", StringComparison.OrdinalIgnoreCase) || keyAction.Equals("ESCAPE", StringComparison.OrdinalIgnoreCase))
            {
                inputSimulator.Keyboard.KeyUp(VirtualKeyCode.ESCAPE);
            }
        }
        
        /// <summary>
        /// 문자에 해당하는 가상 키 코드를 반환합니다.
        /// </summary>
        /// <param name="c">변환할 문자</param>
        /// <returns>가상 키 코드</returns>
        private VirtualKeyCode GetVirtualKeyCodeFromChar(char c)
        {
            // 알파벳
            if (c >= 'a' && c <= 'z')
            {
                return (VirtualKeyCode)((int)VirtualKeyCode.VK_A + (c - 'a'));
            }
            if (c >= 'A' && c <= 'Z')
            {
                return (VirtualKeyCode)((int)VirtualKeyCode.VK_A + (c - 'A'));
            }
            // 숫자
            if (c >= '0' && c <= '9')
            {
                return (VirtualKeyCode)((int)VirtualKeyCode.VK_0 + (c - '0'));
            }
            
            // 다른 특수문자는 필요에 따라 추가
            switch (c)
            {
                case ' ': return VirtualKeyCode.SPACE;
                case '\t': return VirtualKeyCode.TAB;
                case '\r': return VirtualKeyCode.RETURN;
                default: return VirtualKeyCode.SPACE; // 기본값
            }
        }
        
        /// <summary>
        /// 매크로 목록을 JSON 파일에서 로드합니다.
        /// </summary>
        private void LoadMacros()
        {
            try
            {
                if (File.Exists(macroFilePath))
                {
                    string json = File.ReadAllText(macroFilePath);
                    macros = JsonSerializer.Deserialize<List<Macro>>(json) ?? new List<Macro>();
                }
                else
                {
                    macros = new List<Macro>();
                }
            }
            catch (Exception)
            {
                macros = new List<Macro>();
            }
        }

        /// <summary>
        /// 매크로 목록을 JSON 파일로 저장합니다.
        /// </summary>
        private void SaveMacros()
        {
            try
            {
                string json = JsonSerializer.Serialize(macros, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(macroFilePath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"매크로 저장 중 오류 발생: {ex.Message}");
            }
        }

        /// <summary>
        /// 매크로 실행 이벤트를 발생시킵니다.
        /// </summary>
        /// <param name="keyword">실행된 키워드</param>
        /// <param name="keyAction">실행된 키보드 동작</param>
        /// <param name="success">실행 성공 여부</param>
        /// <param name="errorMessage">오류 메시지 (실패 시)</param>
        protected virtual void OnMacroExecuted(string keyword, string keyAction, bool success, string? errorMessage = null)
        {
            MacroExecuted?.Invoke(this, new MacroExecutionEventArgs(keyword, keyAction, success, errorMessage));
        }

        /// <summary>
        /// 상태 변경 이벤트를 발생시킵니다.
        /// </summary>
        /// <param name="status">상태 메시지</param>
        protected virtual void OnStatusChanged(string status)
        {
            StatusChanged?.Invoke(this, status);
        }

        /// <summary>
        /// 모든 매크로 목록을 반환합니다.
        /// </summary>
        /// <returns>매크로 목록</returns>
        public List<Macro> GetAllMacros()
        {
            return macros.ToList(); // 원본 목록의 복사본 반환
        }

        /// <summary>
        /// 사용 가능한 모든 프리셋 정보를 반환합니다.
        /// </summary>
        /// <returns>프리셋 정보 목록</returns>
        public List<PresetInfo> GetAvailablePresets()
        {
            try
            {
                List<PresetInfo> presets = new List<PresetInfo>();
                
                if (!Directory.Exists(presetDirectory))
                {
                    return presets;
                }
                
                // 프리셋 디렉토리의 모든 JSON 파일 검색
                foreach (string filePath in Directory.GetFiles(presetDirectory, "*.json"))
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath);
                        FileInfo fileInfo = new FileInfo(filePath);
                        
                        // 매크로 수 계산
                        int macroCount = 0;
                        try
                        {
                            string json = File.ReadAllText(filePath);
                            List<Macro> presetMacros = JsonSerializer.Deserialize<List<Macro>>(json);
                            macroCount = presetMacros?.Count ?? 0;
                        }
                        catch
                        {
                            // 파일 읽기 오류 시 매크로 수를 0으로 설정
                            macroCount = 0;
                        }
                        
                        // 프리셋 정보 생성
                        PresetInfo preset = new PresetInfo
                        {
                            Name = fileName,
                            FilePath = filePath,
                            MacroCount = macroCount,
                            LastModified = fileInfo.LastWriteTime
                        };
                        
                        presets.Add(preset);
                    }
                    catch (Exception ex)
                    {
                        OnStatusChanged($"프리셋 정보 로드 오류: {ex.Message}");
                    }
                }
                
                return presets;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"프리셋 목록 로드 오류: {ex.Message}");
                return new List<PresetInfo>();
            }
        }

        /// <summary>
        /// 외부 프리셋 파일을 가져옵니다.
        /// </summary>
        /// <param name="filePath">가져올 프리셋 파일 경로</param>
        /// <returns>가져오기 성공 여부</returns>
        public bool ImportPreset(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    OnStatusChanged("가져올 프리셋 파일이 존재하지 않습니다.");
                    return false;
                }
                
                // 프리셋 파일 유효성 검사
                try
                {
                    string json = File.ReadAllText(filePath);
                    JsonSerializer.Deserialize<List<Macro>>(json);
                }
                catch
                {
                    OnStatusChanged("유효하지 않은 프리셋 파일입니다.");
                    return false;
                }
                
                // 파일 이름 추출 및 중복 확인
                string fileName = Path.GetFileName(filePath);
                string destFilePath = Path.Combine(presetDirectory, fileName);
                
                if (File.Exists(destFilePath) && 
                    MessageBox.Show("같은 이름의 프리셋이 이미 존재합니다. 덮어쓰시겠습니까?", 
                        "프리셋 가져오기", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return false;
                }
                
                // 프리셋 파일 복사
                File.Copy(filePath, destFilePath, true);
                OnStatusChanged($"프리셋 '{Path.GetFileNameWithoutExtension(fileName)}' 가져오기 완료");
                return true;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"프리셋 가져오기 오류: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 저장된 프리셋을 삭제합니다.
        /// </summary>
        /// <param name="presetName">삭제할 프리셋 이름</param>
        /// <returns>삭제 성공 여부</returns>
        public bool DeletePreset(string presetName)
        {
            try
            {
                string presetFilePath = Path.Combine(presetDirectory, $"{presetName}.json");
                
                if (!File.Exists(presetFilePath))
                {
                    OnStatusChanged("삭제할 프리셋 파일이 존재하지 않습니다.");
                    return false;
                }
                
                File.Delete(presetFilePath);
                OnStatusChanged($"프리셋 '{presetName}' 삭제 완료");
                return true;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"프리셋 삭제 오류: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// 매크로 정보를 담는 클래스입니다.
    /// </summary>
    public class Macro
    {
        /// <summary>
        /// 음성 명령어 키워드입니다.
        /// </summary>
        public string Keyword { get; set; }

        /// <summary>
        /// 실행할 키보드 동작입니다.
        /// </summary>
        public string KeyAction { get; set; }
        
        /// <summary>
        /// 매크로 액션 유형입니다.
        /// </summary>
        public MacroActionType ActionType { get; set; } = MacroActionType.Default;
        
        /// <summary>
        /// 매크로 액션 추가 매개변수입니다.
        /// 액션 유형에 따라 다른 용도로 사용됩니다:
        /// - Repeat: 반복 횟수
        /// - Hold: 유지 시간(밀리초)
        /// - Turbo: 연타 속도(밀리초 간격)
        /// - Combo: 키 사이 간격(밀리초)
        /// </summary>
        public int ActionParameters { get; set; } = 0;
    }

    /// <summary>
    /// 프리셋 정보를 담는 클래스입니다.
    /// </summary>
    public class PresetInfo
    {
        /// <summary>
        /// 프리셋 이름입니다.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 프리셋 파일 경로입니다.
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 프리셋에 포함된 매크로 수입니다.
        /// </summary>
        public int MacroCount { get; set; }
        
        /// <summary>
        /// 프리셋 파일이 마지막으로 수정된 날짜입니다.
        /// </summary>
        public DateTime LastModified { get; set; }
    }
} 