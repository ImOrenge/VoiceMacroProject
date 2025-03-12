using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using NAudio.CoreAudioApi;

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
        private float microphoneVolume = 1.0f; // 마이크 볼륨 기본값 (0.0f - 1.0f)
        
        // CoreAudio 관련 객체
        private MMDevice activeMicrophoneDevice;
        private MMDeviceEnumerator deviceEnumerator;
        private bool isCoreAudioAvailable = false;

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
            
            // CoreAudio 초기화 시도
            InitializeCoreAudio();
        }
        
        /// <summary>
        /// CoreAudio API를 초기화합니다.
        /// </summary>
        private void InitializeCoreAudio()
        {
            try
            {
                deviceEnumerator = new MMDeviceEnumerator();
                // 현재 활성화된 마이크 가져오기
                activeMicrophoneDevice = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);
                
                // 초기화에 성공하면 플래그 설정
                isCoreAudioAvailable = activeMicrophoneDevice != null;
                
                if (isCoreAudioAvailable)
                {
                    // 현재 시스템 볼륨을 내부 변수에 반영
                    microphoneVolume = activeMicrophoneDevice.AudioEndpointVolume.MasterVolumeLevelScalar;
                    RecordingStatusChanged?.Invoke(this, "시스템 마이크 연결 성공");
                }
                else
                {
                    RecordingStatusChanged?.Invoke(this, "활성화된 마이크를 찾을 수 없습니다.");
                }
            }
            catch (Exception ex)
            {
                isCoreAudioAvailable = false;
                RecordingStatusChanged?.Invoke(this, $"시스템 마이크 연결 실패: {ex.Message}");
            }
        }

        /// <summary>
        /// 시스템에 사용 가능한 마이크가 있는지 확인합니다.
        /// </summary>
        /// <returns>마이크가 있으면 true, 없으면 false</returns>
        public bool HasMicrophone()
        {
            try
            {
                int deviceCount = WaveInEvent.DeviceCount;
                return deviceCount > 0;
            }
            catch
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
                    try
                    {
                        if (silenceDetectionCts != null && !silenceDetectionCts.IsCancellationRequested)
                        {
                            silenceDetectionCts.Cancel();
                        }
                        StopRecording("사용자가 취소함");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"녹음 취소 중 오류: {ex.Message}");
                        // 취소 중 오류가 발생해도 녹음은 중지해야 함
                        try
                        {
                            StopRecording("취소 처리 중 오류 발생");
                        }
                        catch { }
                    }
                }))
                {
                    // 녹음 제한 시간 설정 (최대 녹음 시간)
                    _ = Task.Delay(settings?.MaximumRecordingDuration ?? 10000, silenceDetectionCts.Token)
                        .ContinueWith(t => 
                        {
                            if (!t.IsCanceled && isRecording)
                            {
                                try
                                {
                                    StopRecording("최대 녹음 시간 도달");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"녹음 중지 중 오류: {ex.Message}");
                                }
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
                        catch (Exception ex)
                        {
                            Console.WriteLine($"오디오 최적화 중 오류: {ex.Message}");
                            // 최적화 실패 시 원본 오디오 반환
                            return audioData;
                        }
                        finally
                        {
                            // 임시 파일 삭제
                            try
                            {
                                if (File.Exists(tempFile))
                                {
                                    File.Delete(tempFile);
                                }
                            }
                            catch (Exception ex) 
                            {
                                Console.WriteLine($"임시 파일 삭제 중 오류: {ex.Message}");
                            }
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
                
                try
                {
                    if (silenceDetectionCts != null)
                    {
                        silenceDetectionCts.Dispose();
                        silenceDetectionCts = null;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"리소스 정리 중 오류: {ex.Message}");
                }
                
                try
                {
                    if (writer != null)
                    {
                        writer.Dispose();
                        writer = null;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WAV 파일 작성기 정리 중 오류: {ex.Message}");
                }
            }
        }

        private async Task<byte[]> OptimizeAudioAsync(byte[] audioData)
        {
            if (audioData == null || audioData.Length == 0)
            {
                return Array.Empty<byte>();
            }
            
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
                    bool soundDetected = dbValue > (settings?.RecordingThresholdDb ?? -40);

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
                        
                        if (silenceDuration.TotalMilliseconds > (settings?.SilenceTimeout ?? 1500))
                        {
                            // 최소 녹음 시간 확인
                            TimeSpan recordingDuration = DateTime.Now - recordingStartTime;
                            
                            if (recordingDuration.TotalMilliseconds >= (settings?.MinimumRecordingDuration ?? 1000))
                            {
                                try
                                {
                                    StopRecording("침묵 감지");
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"녹음 중지 중 오류: {ex.Message}");
                                    break;
                                }
                            }
                        }
                    }

                    // 현재 볼륨 레벨 알림
                    AudioLevelChanged?.Invoke(this, dbValue);
                    
                    try
                    {
                        await Task.Delay(100, cancellationToken); // 100ms 간격으로 체크
                    }
                    catch (OperationCanceledException)
                    {
                        // 취소 처리
                        break;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // 취소 처리
            }
            catch (Exception ex)
            {
                Console.WriteLine($"침묵 감지 오류: {ex.Message}");
                // 오류 발생 시 녹음 중지 시도
                try
                {
                    if (isRecording)
                    {
                        StopRecording("침묵 감지 오류");
                    }
                }
                catch { }
            }
        }

        private void WaveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            try
            {
                if (writer != null && isRecording && e != null && e.Buffer != null && e.BytesRecorded > 0)
                {
                    try
                    {
                        // 오디오 데이터 파일에 저장
                        writer.Write(e.Buffer, 0, e.BytesRecorded);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"오디오 데이터 저장 중 오류: {ex.Message}");
                        // 저장 실패 시 녹음 중지 시도
                        try
                        {
                            StopRecording("데이터 저장 오류");
                        }
                        catch { }
                        return;
                    }
                    
                    // 볼륨 레벨 계산
                    float maxSample = 0f;
                    
                    try
                    {
                        for (int i = 0; i < e.BytesRecorded - 1; i += 2)
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
                    catch (Exception ex)
                    {
                        Console.WriteLine($"오디오 레벨 계산 중 오류: {ex.Message}");
                        // 계산 오류는 무시하고 계속 진행
                        maxSample = 0f;
                    }
                    
                    // 오디오 레벨 이벤트 발생 (마이크 레벨을 실시간으로 UI에 표시하기 위함)
                    AudioLevelChanged?.Invoke(this, maxSample);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"오디오 데이터 처리 중 오류: {ex.Message}");
                // 심각한 오류 발생 시 녹음 중지 시도
                try
                {
                    StopRecording("데이터 처리 오류");
                }
                catch { }
            }
        }

        private void WaveIn_RecordingStopped(object sender, StoppedEventArgs e)
        {
            try
            {
                if (writer != null)
                {
                    try
                    {
                        writer.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"WAV 파일 작성기 정리 중 오류: {ex.Message}");
                    }
                    finally
                    {
                        writer = null;
                    }
                }
                
                if (e?.Exception != null)
                {
                    if (recordingTcs != null && !recordingTcs.Task.IsCompleted)
                    {
                        recordingTcs.TrySetException(e.Exception);
                    }
                }
                else if (recordingTcs != null && !recordingTcs.Task.IsCompleted)
                {
                    recordingTcs.TrySetResult(tempFile ?? string.Empty);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"녹음 중지 처리 중 오류: {ex.Message}");
                
                // 마지막 시도로 TaskCompletionSource 완료 처리
                try
                {
                    if (recordingTcs != null && !recordingTcs.Task.IsCompleted)
                    {
                        recordingTcs.TrySetResult(tempFile ?? string.Empty);
                    }
                }
                catch { }
            }
        }

        private void StopRecording(string reason)
        {
            if (!isRecording)
            {
                return; // 이미 중지된 경우 무시
            }
            
            try
            {
                RecordingStatusChanged?.Invoke(this, $"녹음 중단: {reason}");
                
                try
                {
                    if (waveIn != null)
                    {
                        waveIn.StopRecording();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"녹음 중지 중 오류: {ex.Message}");
                }
                
                isRecording = false;
                
                try
                {
                    if (recordingTcs != null && !recordingTcs.Task.IsCompleted)
                    {
                        recordingTcs.TrySetResult(tempFile ?? string.Empty);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"녹음 완료 처리 중 오류: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"녹음 중지 처리 중 오류: {ex.Message}");
                
                // 상태 강제 변경
                isRecording = false;
                
                // 마지막 시도로 TaskCompletionSource 완료 처리
                try
                {
                    if (recordingTcs != null && !recordingTcs.Task.IsCompleted)
                    {
                        recordingTcs.TrySetResult(tempFile ?? string.Empty);
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 시스템 마이크 볼륨을 설정합니다. (0.0f ~ 1.0f 사이의 값)
        /// </summary>
        /// <param name="volume">볼륨 값 (0.0f ~ 1.0f)</param>
        public void SetMicrophoneVolume(float volume)
        {
            // 범위 제한
            if (volume < 0.0f) volume = 0.0f;
            if (volume > 1.0f) volume = 1.0f;
            
            // 내부 변수 업데이트
            microphoneVolume = volume;
            
            try
            {
                // CoreAudio를 통해 시스템 마이크 볼륨 설정
                if (isCoreAudioAvailable && activeMicrophoneDevice != null)
                {
                    // 디바이스 상태 확인
                    if (activeMicrophoneDevice.State == DeviceState.Active)
                    {
                        // 시스템 마이크 볼륨 설정
                        activeMicrophoneDevice.AudioEndpointVolume.MasterVolumeLevelScalar = volume;
                        RecordingStatusChanged?.Invoke(this, $"시스템 마이크 볼륨이 {volume * 100:F0}%로 설정되었습니다.");
                    }
                    else
                    {
                        // 디바이스가 비활성화된 경우 재연결 시도
                        RefreshAudioDevices();
                        if (isCoreAudioAvailable && activeMicrophoneDevice != null && 
                            activeMicrophoneDevice.State == DeviceState.Active)
                        {
                            activeMicrophoneDevice.AudioEndpointVolume.MasterVolumeLevelScalar = volume;
                            RecordingStatusChanged?.Invoke(this, $"마이크 연결 재시도 후 볼륨이 {volume * 100:F0}%로 설정되었습니다.");
                        }
                        else
                        {
                            RecordingStatusChanged?.Invoke(this, "마이크가 비활성화되어 볼륨을 설정할 수 없습니다.");
                        }
                    }
                }
                else
                {
                    // CoreAudio를 사용할 수 없는 경우, 내부 변수만 업데이트
                    RecordingStatusChanged?.Invoke(this, $"내부 마이크 볼륨이 {volume * 100:F0}%로 설정되었습니다. (시스템 연동 없음)");
                    
                    // CoreAudio 재초기화 시도
                    if (!isCoreAudioAvailable)
                    {
                        InitializeCoreAudio();
                    }
                }
            }
            catch (Exception ex)
            {
                // 에러 로깅 및 알림
                RecordingStatusChanged?.Invoke(this, $"마이크 볼륨 설정 오류: {ex.Message}");
                
                // 에러 발생 시 CoreAudio 재초기화 시도
                try
                {
                    isCoreAudioAvailable = false;
                    activeMicrophoneDevice?.Dispose();
                    activeMicrophoneDevice = null;
                    deviceEnumerator?.Dispose();
                    deviceEnumerator = null;
                    
                    InitializeCoreAudio();
                }
                catch
                {
                    // 재초기화 실패 시 무시
                }
            }
        }

        /// <summary>
        /// 현재 마이크 볼륨을 반환합니다.
        /// </summary>
        /// <returns>볼륨 값 (0.0f ~ 1.0f)</returns>
        public float GetMicrophoneVolume()
        {
            try
            {
                // CoreAudio를 통해 시스템 마이크 볼륨 가져오기
                if (isCoreAudioAvailable && activeMicrophoneDevice != null && 
                    activeMicrophoneDevice.State == DeviceState.Active)
                {
                    // 시스템 볼륨으로 내부 변수 업데이트
                    microphoneVolume = activeMicrophoneDevice.AudioEndpointVolume.MasterVolumeLevelScalar;
                }
            }
            catch (Exception ex)
            {
                RecordingStatusChanged?.Invoke(this, $"마이크 볼륨 확인 오류: {ex.Message}");
            }
            
            return microphoneVolume;
        }
        
        /// <summary>
        /// 오디오 디바이스 목록을 새로고침합니다.
        /// </summary>
        public void RefreshAudioDevices()
        {
            try
            {
                // 기존 리소스 정리
                activeMicrophoneDevice?.Dispose();
                activeMicrophoneDevice = null;
                deviceEnumerator?.Dispose();
                deviceEnumerator = null;
                
                // CoreAudio 다시 초기화
                InitializeCoreAudio();
            }
            catch (Exception ex)
            {
                isCoreAudioAvailable = false;
                RecordingStatusChanged?.Invoke(this, $"오디오 디바이스 새로고침 실패: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 사용 가능한 마이크 목록을 가져옵니다.
        /// </summary>
        /// <returns>마이크 이름 목록</returns>
        public List<string> GetAvailableMicrophones()
        {
            List<string> microphoneList = new List<string>();
            
            try
            {
                // WaveIn 디바이스 목록 가져오기
                for (int i = 0; i < WaveInEvent.DeviceCount; i++)
                {
                    var capabilities = WaveInEvent.GetCapabilities(i);
                    microphoneList.Add(capabilities.ProductName);
                }
            }
            catch (Exception ex)
            {
                RecordingStatusChanged?.Invoke(this, $"마이크 목록 가져오기 실패: {ex.Message}");
            }
            
            return microphoneList;
        }

        /// <summary>
        /// 리소스를 해제합니다.
        /// </summary>
        public void Dispose()
        {
            try
            {
                waveIn?.Dispose();
                writer?.Dispose();
                activeMicrophoneDevice?.Dispose();
                deviceEnumerator?.Dispose();
            }
            catch
            {
                // 리소스 해제 중 오류 무시
            }
        }
    }
} 