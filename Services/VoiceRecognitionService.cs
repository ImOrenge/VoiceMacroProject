using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NAudio.Wave;
using Whisper.net;
using Whisper.net.Ggml;
using System.Speech.Recognition;
using System.Net.Http;

namespace VoiceMacro.Services
{
    public class VoiceRecognitionService : IDisposable
    {
        private readonly MacroService macroService;
        private readonly AppSettings settings;
        private readonly OpenAIService? openAIService;
        private readonly AudioRecordingService audioRecorder;
        private readonly WhisperProcessor localWhisperProcessor;
        private CancellationTokenSource? cancellationTokenSource;
        private bool isListening = false;
        private bool isInitialized = false;
        private readonly string modelPath;
        private SpeechRecognitionEngine? recognizer;

        public event EventHandler<string>? SpeechRecognized;
        public event EventHandler<string>? StatusChanged;
        public event EventHandler<float>? AudioLevelChanged;

        public VoiceRecognitionService(MacroService macroService, AppSettings? settings = null, OpenAIService? openAIService = null)
        {
            this.macroService = macroService;
            this.settings = settings ?? new AppSettings();
            this.openAIService = openAIService;
            this.modelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "whisper-model.bin");
            this.audioRecorder = new AudioRecordingService();
            this.localWhisperProcessor = new WhisperProcessor();
            
            // 오디오 레벨 변경 이벤트 연결
            this.audioRecorder.AudioLevelChanged += (s, level) => AudioLevelChanged?.Invoke(this, level);
            
            InitializeAsync();
            InitializeRecognizer(); // System.Speech 음성 인식 엔진 초기화
        }

        private async void InitializeAsync()
        {
            try
            {
                // 모델 파일이 없으면 다운로드
                if (!File.Exists(modelPath))
                {
                    OnStatusChanged("음성 인식을 위한 모델 파일 다운로드 중...");
                    await DownloadModelAsync();
                    OnStatusChanged("모델 파일 다운로드 완료");
                }

                isInitialized = true;
                OnStatusChanged("음성 인식 엔진 초기화 완료");
            }
            catch (Exception ex)
            {
                OnStatusChanged($"음성 인식 엔진 초기화 실패: {ex.Message}");
            }
        }

        private async Task DownloadModelAsync()
        {
            // 기본 Whisper 모델 다운로드
            using var modelStream = await WhisperGgmlDownloader.GetGgmlModelAsync(GgmlType.Tiny);
            using var fileStream = File.Create(modelPath);
            await modelStream.CopyToAsync(fileStream);
        }

        private void InitializeRecognizer()
        {
            try
            {
                if (recognizer != null)
                {
                    // 기존 recognizer가 있으면 정리
                    recognizer.SpeechRecognized -= Recognizer_SpeechRecognized;
                    recognizer.Dispose();
                }

                recognizer = new SpeechRecognitionEngine();
                var grammar = new DictationGrammar();
                recognizer.LoadGrammar(grammar);

                recognizer.SpeechRecognized += Recognizer_SpeechRecognized;
                recognizer.SetInputToDefaultAudioDevice();
                
                OnStatusChanged("시스템 음성 인식 엔진 초기화 완료");
            }
            catch (Exception ex)
            {
                OnStatusChanged($"시스템 음성 인식 엔진 초기화 실패: {ex.Message}");
            }
        }

        private void Recognizer_SpeechRecognized(object? sender, SpeechRecognizedEventArgs e)
        {
            string recognizedText = e.Result.Text;
            OnSpeechRecognized(recognizedText);
            
            // 매크로 처리
            macroService.ProcessVoiceCommand(recognizedText);
        }

        public async Task StartListening()
        {
            if (isListening)
                return;

            if (!isInitialized)
            {
                OnStatusChanged("음성 인식 엔진이 초기화되지 않았습니다.");
                return;
            }

            // 마이크 확인
            if (!HasMicrophone())
            {
                OnStatusChanged("연결된 마이크가 없습니다. 마이크를 연결한 후 다시 시도하세요.");
                return;
            }

            try
            {
                cancellationTokenSource = new CancellationTokenSource();
                isListening = true;
                OnStatusChanged("음성 인식 대기 중...");
                
                // 백그라운드에서 음성 인식 처리 시작
                _ = ListenContinuouslyAsync(cancellationTokenSource.Token);
            }
            catch (Exception ex)
            {
                OnStatusChanged($"음성 인식 시작 오류: {ex.Message}");
                isListening = false;
            }
        }

        public void StopListening()
        {
            if (!isListening)
                return;

            try
            {
                cancellationTokenSource?.Cancel();
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
                isListening = false;
                OnStatusChanged("음성 인식 중지됨");
            }
            catch (Exception ex)
            {
                OnStatusChanged($"음성 인식 중지 오류: {ex.Message}");
            }
        }

        private async Task ListenContinuouslyAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested && isListening)
            {
                try
                {
                    // 설정에 따라 음성 인식 방식 선택
                    string? recognizedText = null;
                    
                    if (settings.UseLocalVoiceRecognition)
                    {
                        // 로컬 Whisper 모델 사용
                        byte[] audioData = await audioRecorder.RecordSpeechAsync(cancellationToken);
                        recognizedText = await localWhisperProcessor.ProcessAudioAsync(audioData, settings.WhisperLanguage, cancellationToken);
                    }
                    else if (settings.UseOpenAI && openAIService != null)
                    {
                        // OpenAI API 사용
                        byte[] audioData = await audioRecorder.RecordSpeechAsync(cancellationToken);
                        recognizedText = await openAIService.TranscribeAudioAsync(audioData, settings.WhisperLanguage, cancellationToken);
                    }
                    else
                    {
                        // 시스템 음성 인식 사용 (System.Speech)
                        await Task.Delay(100, cancellationToken); // CPU 부하 방지용 지연
                        continue; // 음성 인식은 이벤트 기반으로 처리됨
                    }
                    
                    if (!string.IsNullOrEmpty(recognizedText))
                    {
                        OnSpeechRecognized(recognizedText);
                        macroService.ProcessVoiceCommand(recognizedText);
                    }
                }
                catch (OperationCanceledException)
                {
                    // 취소 요청에 의한 종료는 정상 처리
                    break;
                }
                catch (Exception ex)
                {
                    OnStatusChanged($"음성 인식 오류: {ex.Message}");
                    await Task.Delay(1000, cancellationToken); // 오류 발생 시 잠시 대기
                }
            }
        }

        public void UpdateSettings(AppSettings? newSettings)
        {
            // 설정 업데이트
            if (newSettings != null)
            {
                settings.CopyFrom(newSettings);
                
                if (isListening)
                {
                    // 음성 인식 재시작
                    StopListening();
                    _ = StartListening();
                }
            }
        }

        public void Dispose()
        {
            // 리소스 정리
            StopListening();
            cancellationTokenSource?.Dispose();
            localWhisperProcessor?.Dispose();
            audioRecorder?.Dispose();
            recognizer?.Dispose();
        }

        public async Task<string> RecognizeAudioAsync(byte[] audioData, string language = "ko")
        {
            try
            {
                if (settings.UseOpenAI && openAIService != null)
                {
                    // OpenAI API로 음성 인식
                    return await openAIService.TranscribeAudioAsync(audioData, language);
                }
                else
                {
                    // 로컬 Whisper 모델로 음성 인식
                    return await localWhisperProcessor.ProcessAudioAsync(audioData, language);
                }
            }
            catch (Exception ex)
            {
                OnStatusChanged($"음성 인식 처리 오류: {ex.Message}");
                return string.Empty;
            }
        }

        public bool HasMicrophone()
        {
            return WaveInEvent.DeviceCount > 0;
        }

        private class WhisperProcessor : IDisposable
        {
            private WhisperFactory? factory;
            private object? processor; // 동적 타입으로 처리

            public WhisperProcessor()
            {
                // 사용 시점에 초기화
            }

            public Task<string> ProcessAudioAsync(byte[] audioData, string language = "ko", CancellationToken cancellationToken = default)
            {
                // 임시 구현 - 실제 Whisper 모델을 사용한 처리는 추후 구현
                return Task.FromResult($"[Whisper 인식 결과 - {language}]");
            }

            public void Dispose()
            {
                if (processor is IDisposable disposable)
                {
                    disposable.Dispose();
                }
                factory?.Dispose();
            }
        }

        // 음성 인식 이벤트 발생
        protected virtual void OnSpeechRecognized(string recognizedText)
        {
            SpeechRecognized?.Invoke(this, recognizedText);
        }

        // 상태 변경 이벤트 발생
        protected virtual void OnStatusChanged(string status)
        {
            StatusChanged?.Invoke(this, status);
        }
    }
} 