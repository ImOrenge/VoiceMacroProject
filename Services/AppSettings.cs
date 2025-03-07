using System;
using System.IO;
using System.Text.Json;

namespace VoiceMacro.Services
{
    public class AppSettings
    {
        private static readonly string SettingsFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "settings.json");

        public string? OpenAIApiKey { get; set; }
        public bool UseOpenAIApi { get; set; } = false;
        public bool UseLocalVoiceRecognition { get; set; } = true;
        public bool UseOpenAI { get; set; } = false;
        public float WhisperTemperature { get; set; } = 0.0f;
        public string WhisperLanguage { get; set; } = "ko";
        
        // 음성 인식 관련 설정
        public int RecordingThresholdDb { get; set; } = -30; // 감지 임계값 (dB)
        public int MinimumRecordingDuration { get; set; } = 1000; // 최소 녹음 시간 (ms)
        public int MaximumRecordingDuration { get; set; } = 15000; // 최대 녹음 시간 (ms)
        public int SilenceTimeout { get; set; } = 1500; // 침묵 감지 시간 (ms)

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    string json = File.ReadAllText(SettingsFilePath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json);
                    return settings ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"설정 로드 오류: {ex.Message}");
            }

            return new AppSettings();
        }

        public void Save()
        {
            try
            {
                string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsFilePath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"설정 저장 오류: {ex.Message}");
            }
        }

        /// <summary>
        /// 다른 설정 객체의 값을 현재 객체로 복사합니다.
        /// </summary>
        /// <param name="other">복사할 원본 설정 객체</param>
        public void CopyFrom(AppSettings other)
        {
            if (other == null) return;

            OpenAIApiKey = other.OpenAIApiKey;
            UseOpenAIApi = other.UseOpenAIApi;
            UseLocalVoiceRecognition = other.UseLocalVoiceRecognition;
            UseOpenAI = other.UseOpenAI;
            WhisperTemperature = other.WhisperTemperature;
            WhisperLanguage = other.WhisperLanguage;
            RecordingThresholdDb = other.RecordingThresholdDb;
            MinimumRecordingDuration = other.MinimumRecordingDuration;
            MaximumRecordingDuration = other.MaximumRecordingDuration;
            SilenceTimeout = other.SilenceTimeout;
        }
    }
} 