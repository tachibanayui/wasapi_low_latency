# WASAPI Low Latency

This project demonstrates low-latency audio processing using the Windows Audio Session API (WASAPI) in Rust. It provides functionality for capturing and rendering audio with minimal latency, supporting both device-based and process-based audio input. Low latency audio is achieved by using IAudioClient3 to request smaller buffers (down to 2.67ms on my local machine) and process captures by using the new
windows 10 api `ActivateAudioInterfaceAsync` with `VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK`

## Features

-   **Low-latency audio processing**: Utilizes WASAPI IAudioClient3 for efficient audio capture and rendering.
-   **Device and process input**: Supports capturing audio from specific devices or processes.
-   **MMCSS Pro Audio task registration**: Optimizes thread priority for audio processing.

## Automatically fill stdin

While this project ask prompt user for input interactively, you can create a `stdio.txt` file to automatically fill in prompts. This is convinient for debugging

## License

This project is licensed under the MIT License.
