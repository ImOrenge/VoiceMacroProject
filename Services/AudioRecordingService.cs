using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

namespace VoiceMacro.Services
{
    public class AudioRecordingService : IDisposable
    {
        private readonly WaveInEvent waveIn;
        private WaveFileWriter writer;
        private readonly int sampleRate;
        private readonly int channels;
        private string tempFile;
        private TaskCompletionSource<string> recordingTcs;
        private CancellationTokenSource silenceDetectionCts;
        private DateTime lastSoundTime;
        private bool isRecording = false;
        private readonly AppSettings settings;
        private float maxVolume = 0f;
        private readonly object lockObject = new object();

        public event EventHandler<float> AudioLevelChanged;
        public event EventHandler<string> RecordingStatusChanged;

        public AudioRecordingService(AppSettings? settings = null)
        {
            this.settings = settings ?? new AppSettings();
            sampleRate = 16000; // Whisper에 최적화된 샘플 레이트
            channels = 1; // 모노

            waveIn = new WaveInEvent
            {
                WaveFormat = new WaveFormat(sampleRate, 16, channels),
                BufferMilliseconds = 50
            };

            waveIn.DataAvailable += WaveIn_DataAvailable;
            waveIn.RecordingStopped += WaveIn_RecordingStopped;
        }

        /// <summary>
        /// 시스템에 사용 가능한 마이크가 있는지 확인합니다.
        /// </summary>
        /// <returns>마이크가 있으면 true, 없으면 false</returns>
        public bool HasMicrophone()
        {
            try
            {
                int deviceCount = WaveIn.DeviceCount;
                return deviceCount > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 음성 녹음을 시작하고, 음성이 감지되었을 때 녹음을 시작하여 
        /// 자동으로 침묵을 감지하면 중단합니다.
        /// </summary>
        public async Task<byte[]> RecordSpeechAsync(CancellationToken cancellationToken = default)
        {
            // 먼저 마이크 확인
            if (!HasMicrophone())
            {
                RecordingStatusChanged?.Invoke(this, "연결된 마이크가 없습니다.");
                return Array.Empty<byte>();
            }

            if (isRecording)
            {
                throw new InvalidOperationException("이미 녹음 중입니다.");
            }

            RecordingStatusChanged?.Invoke(this, "음성 감지 대기 중...");
            
            try
            {
                tempFile = Path.Combine(Path.GetTempPath(), $"recording_{Guid.NewGuid()}.wav");
                writer = new WaveFileWriter(tempFile, waveIn.WaveFormat);
                
                recordingTcs = new TaskCompletionSource<string>();
                silenceDetectionCts = new CancellationTokenSource();
                
                maxVolume = 0f;
                lastSoundTime = DateTime.Now;
                isRecording = true;

                // 녹음 시작
                waveIn.StartRecording();

                // 취소 토큰 등록
                using (cancellationToken.Register(() => 
                {
                    silenceDetectionCts.Cancel();
                    StopRecording("사용자가 취소함");
                }))
                {
                    // 녹음 제한 시간 설정 (최대 녹음 시간)
                    _ = Task.Delay(settings.MaximumRecordingDuration, silenceDetectionCts.Token)
                        .ContinueWith(t => 
                        {
                            if (!t.IsCanceled && isRecording)
                            {
                                StopRecording("최대 녹음 시간 도달");
                            }
                        }, TaskScheduler.Default);

                    // 침묵 감지 로직 시작
                    _ = DetectSilenceAsync(silenceDetectionCts.Token);

                    // 녹음 완료 대기
                    await recordingTcs.Task;

                    // 녹음 파일 읽기
                    if (File.Exists(tempFile))
                    {
                        byte[] audioData = File.ReadAllBytes(tempFile);
                        
                        try
                        {
                            // WAV 파일 최적화 및 노이즈 제거
                            return await OptimizeAudioAsync(audioData);
                        }
                        finally
                        {
                            // 임시 파일 삭제
                            try
                            {
                                File.Delete(tempFile);
                            }
                            catch { }
                        }
                    }

                    return Array.Empty<byte>();
                }
            }
            catch (Exception ex)
            {
                RecordingStatusChanged?.Invoke(this, $"녹음 오류: {ex.Message}");
                return Array.Empty<byte>();
            }
            finally
            {
                isRecording = false;
                silenceDetectionCts?.Dispose();
            }
        }

        private async Task<byte[]> OptimizeAudioAsync(byte[] audioData)
        {
            try
            {
                // 향후 노이즈 제거 등의 고급 처리를 위한 자리 표시자
                // 지금은 단순히 원본 데이터 반환

                return audioData;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오디오 최적화 오류: {ex.Message}");
                return audioData; // 오류 시 원본 데이터 반환
            }
        }

        private async Task DetectSilenceAsync(CancellationToken cancellationToken)
        {
            try
            {
                bool activeRecording = false;
                DateTime recordingStartTime = DateTime.MinValue;

                while (!cancellationToken.IsCancellationRequested && isRecording)
                {
                    // 현재 최대 볼륨 확인
                    float currentMaxVolume;
                    lock (lockObject)
                    {
                        currentMaxVolume = maxVolume;
                        maxVolume = 0f; // 볼륨 리셋
                    }

                    // 데시벨로 변환 (로그 스케일)
                    float dbValue = 20 * (float)Math.Log10(currentMaxVolume > 0 ? currentMaxVolume : 0.0000001);

                    // 소리 감지
                    bool soundDetected = dbValue > settings.RecordingThresholdDb;

                    if (soundDetected)
                    {
                        lastSoundTime = DateTime.Now;
                        
                        if (!activeRecording)
                        {
                            activeRecording = true;
                            recordingStartTime = DateTime.Now;
                            RecordingStatusChanged?.Invoke(this, "녹음 중...");
                        }
                    }
                    else if (activeRecording)
                    {
                        // 침묵 지속 시간 확인
                        TimeSpan silenceDuration = DateTime.Now - lastSoundTime;
                        
                        if (silenceDuration.TotalMilliseconds > settings.SilenceTimeout)
                        {
                            // 최소 녹음 시간 확인
                            TimeSpan recordingDuration = DateTime.Now - recordingStartTime;
                            
                            if (recordingDuration.TotalMilliseconds >= settings.MinimumRecordingDuration)
                            {
                                StopRecording("침묵 감지");
                                break;
                            }
                        }
                    }

                    // 현재 볼륨 레벨 알림
                    AudioLevelChanged?.Invoke(this, dbValue);
                    
                    await Task.Delay(100, cancellationToken); // 100ms 간격으로 체크
                }
            }
            catch (OperationCanceledException)
            {
                // 취소 처리
            }
            catch (Exception ex)
            {
                Console.WriteLine($"침묵 감지 오류: {ex.Message}");
            }
        }

        private void WaveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (writer != null && isRecording)
            {
                // 오디오 데이터 파일에 저장
                writer.Write(e.Buffer, 0, e.BytesRecorded);
                
                // 볼륨 레벨 계산
                float maxSample = 0f;
                
                for (int i = 0; i < e.BytesRecorded; i += 2)
                {
                    short sample = (short)((e.Buffer[i + 1] << 8) | e.Buffer[i]);
                    float sampleFloat = Math.Abs(sample) / 32768f;
                    maxSample = Math.Max(maxSample, sampleFloat);
                }
                
                lock (lockObject)
                {
                    maxVolume = Math.Max(maxVolume, maxSample);
                }
            }
        }

        private void WaveIn_RecordingStopped(object sender, StoppedEventArgs e)
        {
            if (writer != null)
            {
                writer.Dispose();
                writer = null;
            }
            
            if (e.Exception != null)
            {
                recordingTcs?.TrySetException(e.Exception);
            }
        }

        private void StopRecording(string reason)
        {
            if (isRecording)
            {
                RecordingStatusChanged?.Invoke(this, $"녹음 중단: {reason}");
                
                waveIn.StopRecording();
                isRecording = false;
                
                recordingTcs?.TrySetResult(tempFile);
            }
        }

        public void Dispose()
        {
            silenceDetectionCts?.Cancel();
            waveIn?.StopRecording();
            waveIn?.Dispose();
            writer?.Dispose();
        }
    }
} 