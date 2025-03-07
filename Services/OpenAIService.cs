using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace VoiceMacro.Services
{
    public class OpenAIService
    {
        private readonly string apiKey;
        private readonly HttpClient httpClient;

        public OpenAIService(string apiKey)
        {
            this.apiKey = apiKey;
            httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        }

        /// <summary>
        /// Whisper API를 사용하여 음성을 텍스트로 변환합니다
        /// </summary>
        public async Task<string> TranscribeAudioAsync(byte[] audioData, string language = "ko", CancellationToken cancellationToken = default)
        {
            try
            {
                if (string.IsNullOrEmpty(apiKey))
                {
                    throw new InvalidOperationException("OpenAI API 키가 설정되지 않았습니다.");
                }

                // 임시 파일로 변환
                string tempFile = Path.Combine(Path.GetTempPath(), $"whisper_audio_{Guid.NewGuid()}.wav");
                await File.WriteAllBytesAsync(tempFile, audioData, cancellationToken);

                try
                {
                    // HTTP 직접 요청으로 OpenAI API 호출
                    using var formContent = new MultipartFormDataContent();
                    using var fileStream = new FileStream(tempFile, FileMode.Open);
                    using var fileContent = new StreamContent(fileStream);
                    
                    formContent.Add(fileContent, "file", "audio.wav");
                    formContent.Add(new StringContent(language), "language");
                    formContent.Add(new StringContent("transcribe"), "model");

                    var response = await httpClient.PostAsync("https://api.openai.com/v1/audio/transcriptions", formContent, cancellationToken);
                    response.EnsureSuccessStatusCode();

                    var jsonResponse = await response.Content.ReadAsStringAsync(cancellationToken);
                    var transcription = JsonConvert.DeserializeObject<TranscriptionResponse>(jsonResponse);

                    return transcription?.Text ?? string.Empty;
                }
                finally
                {
                    // 임시 파일 삭제
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"OpenAI 음성 인식 오류: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// API 키가 유효한지 확인합니다 (간단한 API 호출로 테스트)
        /// </summary>
        public async Task<bool> ValidateApiKeyAsync()
        {
            try
            {
                // 간단한 모델 목록 요청으로 키 유효성 테스트
                var response = await httpClient.GetAsync("https://api.openai.com/v1/models");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        private class TranscriptionResponse
        {
            [JsonProperty("text")]
            public string Text { get; set; }
        }
    }
} 