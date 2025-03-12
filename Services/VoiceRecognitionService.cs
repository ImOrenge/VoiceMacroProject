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
                isListening = false;  // 먼저 상태를 변경하여 ListenContinuouslyAsync의 루프가 중단되도록 함
                
                // CancellationTokenSource 정리
                if (cancellationTokenSource != null)
                {
                    try
                    {
                        if (!cancellationTokenSource.IsCancellationRequested)
                        {
                            try
                            {
                                cancellationTokenSource.Cancel();
                            }
                            catch (ObjectDisposedException)
                            {
                                // 이미 Dispose된 경우 무시
                            }
                            catch (Exception ex)
                            {
                                // 다른 예외도 무시하고 로그만 남김
                                System.Diagnostics.Debug.WriteLine($"CancellationToken 취소 중 오류: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // IsCancellationRequested 접근 오류 무시
                    }
                    
                    try
                    {
                        cancellationTokenSource.Dispose();
                    }
                    catch (ObjectDisposedException)
                    {
                        // 이미 Dispose된 경우 무시
                    }
                    catch (Exception ex)
                    {
                        // 다른 예외도 무시하고 로그만 남김
                        System.Diagnostics.Debug.WriteLine($"CancellationTokenSource 해제 중 오류: {ex.Message}");
                    }
                    
                    cancellationTokenSource = null;
                }
                
                OnStatusChanged("음성 인식 중지됨");
            }
            catch (Exception ex)
            {
                OnStatusChanged($"음성 인식 중지 오류: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"StopListening 완료 중 오류: {ex.Message}");
            }
        }

        private async Task ListenContinuouslyAsync(CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested && isListening)
                {
                    try
                    {
                        // 설정에 따라 음성 인식 방식 선택
                        string? recognizedText = null;
                        
                        if (settings.UseLocalVoiceRecognition)
                        {
                            try
                            {
                                // 로컬 Whisper 모델 사용
                                byte[] audioData = await audioRecorder.RecordSpeechAsync(cancellationToken);
                                if (audioData.Length > 0)
                                {
                                    recognizedText = await localWhisperProcessor.ProcessAudioAsync(audioData, settings.WhisperLanguage, cancellationToken);
                                }
                            }
                            catch (OperationCanceledException)
                            {
                                // 작업이 취소된 경우 (정상적인 종료) - 루프 중단
                                break;
                            }
                        }
                        else if (settings.UseOpenAI && openAIService != null)
                        {
                            try
                            {
                                // OpenAI API 사용
                                byte[] audioData = await audioRecorder.RecordSpeechAsync(cancellationToken);
                                if (audioData.Length > 0)
                                {
                                    recognizedText = await openAIService.TranscribeAudioAsync(audioData, settings.WhisperLanguage, cancellationToken);
                                }
                            }
                            catch (OperationCanceledException)
                            {
                                // 작업이 취소된 경우 (정상적인 종료) - 루프 중단
                                break;
                            }
                        }
                        else
                        {
                            // 시스템 음성 인식 사용 (System.Speech)
                            try
                            {
                                await Task.Delay(100, cancellationToken); // CPU 부하 방지용 지연
                            }
                            catch (OperationCanceledException)
                            {
                                // 작업이 취소된 경우 (정상적인 종료) - 루프 중단
                                break;
                            }
                            continue; // 음성 인식은 이벤트 기반으로 처리됨
                        }
                        
                        if (!string.IsNullOrEmpty(recognizedText))
                        {
                            OnSpeechRecognized(recognizedText);
                            macroService.ProcessVoiceCommand(recognizedText);
                        }
                    }
                    catch (Exception ex)
                    {
                        OnStatusChanged($"음성 인식 오류: {ex.Message}");
                        try
                        {
                            await Task.Delay(1000, cancellationToken); // 오류 발생 시 잠시 대기
                        }
                        catch (OperationCanceledException)
                        {
                            // 작업이 취소된 경우 루프 중단
                            break;
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // 작업이 취소된 경우 (정상적인 종료)
                OnStatusChanged("음성 인식이 취소되었습니다.");
            }
            catch (Exception ex)
            {
                OnStatusChanged($"음성 인식 중 오류: {ex.Message}");
            }
            finally
            {
                // 리스닝 상태 정리
                isListening = false;
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
            try
            {
                // 리소스 정리
                // 먼저 리스닝 중지
                if (isListening)
                {
                    StopListening();
                }
                
                // 각 리소스 정리 중 발생하는 예외를 개별적으로 처리
                try
                {
                    if (cancellationTokenSource != null)
                    {
                        if (!cancellationTokenSource.IsCancellationRequested)
                        {
                            try { cancellationTokenSource.Cancel(); } catch { }
                        }
                        try { cancellationTokenSource.Dispose(); } catch { }
                        cancellationTokenSource = null;
                    }
                }
                catch { }
                
                try { localWhisperProcessor?.Dispose(); } catch { }
                try { audioRecorder?.Dispose(); } catch { }
                try { recognizer?.Dispose(); } catch { }
            }
            catch (Exception ex)
            {
                // 최종적으로 발생하는 모든 예외는 여기서 처리하고 로그만 남김
                System.Diagnostics.Debug.WriteLine($"VoiceRecognitionService 해제 중 오류: {ex.Message}");
            }
        }

        public async Task<string> RecognizeAudioAsync(byte[] audioData, string language = "ko")
        {
            if (audioData == null || audioData.Length == 0)
            {
                OnStatusChanged("인식할 오디오 데이터가 없습니다.");
                return string.Empty;
            }
            
            try
            {
                if (settings?.UseOpenAI == true && openAIService != null)
                {
                    // OpenAI API로 음성 인식
                    string result = await openAIService.TranscribeAudioAsync(audioData, language ?? "ko");
                    return result ?? string.Empty;
                }
                else
                {
                    // 로컬 Whisper 모델로 음성 인식
                    if (localWhisperProcessor != null)
                    {
                        string result = await localWhisperProcessor.ProcessAudioAsync(audioData, language ?? "ko");
                        return result ?? string.Empty;
                    }
                    else
                    {
                        OnStatusChanged("Whisper 프로세서가 초기화되지 않았습니다.");
                        return string.Empty;
                    }
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
            try
            {
                return WaveInEvent.DeviceCount > 0;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"마이크 감지 오류: {ex.Message}");
                return false;
            }
        }

        private class WhisperProcessor : IDisposable
        {
            private WhisperFactory? factory;
            private object? processor; // 동적 타입으로 처리
            private bool isDisposed = false;

            public WhisperProcessor()
            {
                // 사용 시점에 초기화
                try
                {
                    // 초기화 로직 (필요시)
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WhisperProcessor 초기화 오류: {ex.Message}");
                }
            }

            public async Task<string> ProcessAudioAsync(byte[] audioData, string language = "ko", CancellationToken cancellationToken = default)
            {
                if (isDisposed)
                {
                    throw new ObjectDisposedException(nameof(WhisperProcessor));
                }
                
                if (audioData == null || audioData.Length == 0)
                {
                    return string.Empty;
                }
                
                try
                {
                    // 임시 구현 - 실제 Whisper 모델을 사용한 처리는 추후 구현
                    await Task.Delay(100, cancellationToken); // 비동기 작업 시뮬레이션
                    return $"[Whisper 인식 결과 - {language ?? "ko"}]";
                }
                catch (OperationCanceledException)
                {
                    // 작업 취소 처리
                    return string.Empty;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Whisper 처리 오류: {ex.Message}");
                    return string.Empty;
                }
            }

            public void Dispose()
            {
                if (isDisposed)
                {
                    return;
                }
                
                try
                {
                    // 리소스 정리
                    if (processor != null)
                    {
                        // processor가 IDisposable을 구현한다면 Dispose 호출
                        if (processor is IDisposable disposable)
                        {
                            disposable.Dispose();
                        }
                        processor = null;
                    }
                    
                    if (factory != null)
                    {
                        // factory가 IDisposable을 구현한다면 Dispose 호출
                        if (factory is IDisposable disposable)
                        {
                            disposable.Dispose();
                        }
                        factory = null;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WhisperProcessor 정리 중 오류: {ex.Message}");
                }
                finally
                {
                    isDisposed = true;
                }
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

        /// <summary>
        /// AudioRecordingService 인스턴스를 반환합니다.
        /// </summary>
        /// <returns>AudioRecordingService 인스턴스</returns>
        public AudioRecordingService GetAudioRecordingService()
        {
            return audioRecorder;
        }
    }
} 