using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;

namespace VoiceMacro.Services
{
    /// <summary>
    /// 매크로 실행 이벤트 인자
    /// </summary>
    public class MacroExecutionEventArgs : EventArgs
    {
        public string Keyword { get; }
        public string KeyAction { get; }
        public bool Success { get; }
        public string? ErrorMessage { get; }

        public MacroExecutionEventArgs(string keyword, string keyAction, bool success, string? errorMessage = null)
        {
            Keyword = keyword;
            KeyAction = keyAction;
            Success = success;
            ErrorMessage = errorMessage;
        }
    }

    public class MacroService
    {
        private List<Macro> macros;
        private readonly string macroFilePath;
        private readonly string presetFolderPath;
        private readonly InputSimulator inputSimulator;

        // 매크로 실행 이벤트
        public event EventHandler<MacroExecutionEventArgs> MacroExecuted;
        
        // 상태 변경 이벤트
        public event EventHandler<string> StatusChanged;

        public MacroService()
        {
            string appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                "VoiceMacro");
            
            // 애플리케이션 데이터 폴더 생성
            if (!Directory.Exists(appDataPath))
            {
                Directory.CreateDirectory(appDataPath);
            }
            
            macroFilePath = Path.Combine(appDataPath, "macros.json");
            presetFolderPath = Path.Combine(appDataPath, "Presets");
            
            // 프리셋 폴더 생성
            if (!Directory.Exists(presetFolderPath))
            {
                Directory.CreateDirectory(presetFolderPath);
            }
            
            inputSimulator = new InputSimulator();
            LoadMacros();
        }

        public void AddMacro(string keyword, string keyAction)
        {
            macros.Add(new Macro { Keyword = keyword, KeyAction = keyAction });
            SaveMacros();
            OnStatusChanged($"매크로 추가됨: {keyword}");
        }

        public void RemoveMacro(string keyword)
        {
            macros.RemoveAll(m => m.Keyword.Equals(keyword, StringComparison.OrdinalIgnoreCase));
            SaveMacros();
            OnStatusChanged($"매크로 삭제됨: {keyword}");
        }

        public List<Macro> GetAllMacros()
        {
            return macros;
        }

        /// <summary>
        /// 현재 매크로 설정을 프리셋으로 저장합니다.
        /// </summary>
        /// <param name="presetName">프리셋 이름</param>
        /// <returns>저장 성공 여부</returns>
        public bool SavePreset(string presetName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(presetName))
                {
                    return false;
                }
                
                // 파일명에 유효하지 않은 문자 제거
                foreach (char invalidChar in Path.GetInvalidFileNameChars())
                {
                    presetName = presetName.Replace(invalidChar, '_');
                }
                
                string presetFilePath = Path.Combine(presetFolderPath, $"{presetName}.json");
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
        /// 프리셋 파일을 불러와 현재 매크로 설정을 교체합니다.
        /// </summary>
        /// <param name="presetFilePath">프리셋 파일 경로</param>
        /// <returns>불러오기 성공 여부</returns>
        public bool LoadPreset(string presetFilePath)
        {
            try
            {
                if (!File.Exists(presetFilePath))
                {
                    OnStatusChanged("프리셋 파일이 존재하지 않습니다.");
                    return false;
                }
                
                string json = File.ReadAllText(presetFilePath);
                var loadedMacros = JsonSerializer.Deserialize<List<Macro>>(json);
                
                if (loadedMacros == null || loadedMacros.Count == 0)
                {
                    OnStatusChanged("프리셋에 매크로가 없습니다.");
                    return false;
                }
                
                // 현재 매크로 설정을 교체
                macros = loadedMacros;
                SaveMacros(); // 현재 매크로 파일에도 저장
                
                string presetName = Path.GetFileNameWithoutExtension(presetFilePath);
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
        /// 프리셋 파일을 불러와 현재 매크로 설정에 추가합니다.
        /// </summary>
        /// <param name="presetFilePath">프리셋 파일 경로</param>
        /// <returns>추가 성공 여부</returns>
        public bool ImportPreset(string presetFilePath)
        {
            try
            {
                if (!File.Exists(presetFilePath))
                {
                    OnStatusChanged("프리셋 파일이 존재하지 않습니다.");
                    return false;
                }
                
                string json = File.ReadAllText(presetFilePath);
                var loadedMacros = JsonSerializer.Deserialize<List<Macro>>(json);
                
                if (loadedMacros == null || loadedMacros.Count == 0)
                {
                    OnStatusChanged("프리셋에 매크로가 없습니다.");
                    return false;
                }
                
                int addedCount = 0;
                
                // 중복 매크로 확인 및 추가
                foreach (var macro in loadedMacros)
                {
                    // 같은 키워드의 매크로가 있는지 확인
                    bool isDuplicate = macros.Exists(m => 
                        m.Keyword.Equals(macro.Keyword, StringComparison.OrdinalIgnoreCase));
                    
                    if (!isDuplicate)
                    {
                        macros.Add(macro);
                        addedCount++;
                    }
                }
                
                if (addedCount > 0)
                {
                    SaveMacros(); // 변경된 매크로 저장
                }
                
                string presetName = Path.GetFileNameWithoutExtension(presetFilePath);
                OnStatusChanged($"프리셋 '{presetName}'에서 {addedCount}개 매크로 추가 완료");
                return addedCount > 0;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"프리셋 가져오기 오류: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// 사용 가능한 모든 프리셋 파일 목록을 반환합니다.
        /// </summary>
        /// <returns>프리셋 파일 정보 목록</returns>
        public List<PresetInfo> GetAvailablePresets()
        {
            try
            {
                List<PresetInfo> presets = new List<PresetInfo>();
                
                if (!Directory.Exists(presetFolderPath))
                {
                    return presets;
                }
                
                string[] presetFiles = Directory.GetFiles(presetFolderPath, "*.json");
                
                foreach (string filePath in presetFiles)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(filePath);
                        string presetName = Path.GetFileNameWithoutExtension(filePath);
                        
                        // 매크로 개수 확인
                        string json = File.ReadAllText(filePath);
                        var presetMacros = JsonSerializer.Deserialize<List<Macro>>(json);
                        int macroCount = presetMacros?.Count ?? 0;
                        
                        presets.Add(new PresetInfo
                        {
                            Name = presetName,
                            FilePath = filePath,
                            LastModified = fileInfo.LastWriteTime,
                            MacroCount = macroCount
                        });
                    }
                    catch
                    {
                        // 개별 프리셋 파일 처리 중 오류는 무시하고 계속 진행
                        continue;
                    }
                }
                
                // 수정 날짜 기준 내림차순 정렬
                presets.Sort((a, b) => b.LastModified.CompareTo(a.LastModified));
                return presets;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"프리셋 목록 조회 오류: {ex.Message}");
                return new List<PresetInfo>();
            }
        }
        
        /// <summary>
        /// 프리셋 파일을 삭제합니다.
        /// </summary>
        /// <param name="presetFilePath">삭제할 프리셋 파일 경로</param>
        /// <returns>삭제 성공 여부</returns>
        public bool DeletePreset(string presetFilePath)
        {
            try
            {
                if (!File.Exists(presetFilePath))
                {
                    return false;
                }
                
                string presetName = Path.GetFileNameWithoutExtension(presetFilePath);
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

        // 음성 명령 처리
        public bool ProcessVoiceCommand(string command)
        {
            if (string.IsNullOrEmpty(command))
            {
                OnStatusChanged("빈 명령어가 입력되었습니다.");
                return false;
            }

            OnStatusChanged($"명령어 처리 중: {command}");
            
            // 기존 ExecuteMacroIfMatched 메서드의 로직을 사용
            return ExecuteMacroIfMatched(command);
        }

        /// <summary>
        /// 인식된 음성에서 매크로를 찾아 실행합니다.
        /// </summary>
        public bool ExecuteMacroIfMatched(string recognizedText)
        {
            if (string.IsNullOrWhiteSpace(recognizedText))
                return false;

            OnStatusChanged($"매크로 검색: {recognizedText}");
            
            foreach (var macro in macros)
            {
                if (recognizedText.Contains(macro.Keyword, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        ExecuteKeyAction(macro.KeyAction);
                        OnMacroExecuted(macro.Keyword, macro.KeyAction, true);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        OnMacroExecuted(macro.Keyword, macro.KeyAction, false, ex.Message);
                        return false;
                    }
                }
            }
            
            OnStatusChanged("일치하는 매크로 없음");
            return false;
        }

        private void ExecuteKeyAction(string keyAction)
        {
            try
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
            catch (Exception ex)
            {
                Console.WriteLine($"키 동작 실행 중 오류 발생: {ex.Message}");
                throw; // 예외를 상위로 전파하여 ProcessVoiceCommand에서 처리하도록 함
            }
        }

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

        // 매크로 실행 이벤트 발생
        protected virtual void OnMacroExecuted(string keyword, string keyAction, bool success, string? errorMessage = null)
        {
            MacroExecuted?.Invoke(this, new MacroExecutionEventArgs(keyword, keyAction, success, errorMessage));
        }

        // 상태 변경 이벤트 발생
        protected virtual void OnStatusChanged(string status)
        {
            StatusChanged?.Invoke(this, status);
        }
    }

    public class Macro
    {
        public string Keyword { get; set; } = string.Empty;
        public string KeyAction { get; set; } = string.Empty;
    }
    
    public class PresetInfo
    {
        public string Name { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public DateTime LastModified { get; set; }
        public int MacroCount { get; set; }
        
        public override string ToString()
        {
            return $"{Name} ({MacroCount}개 매크로)";
        }
    }
} 