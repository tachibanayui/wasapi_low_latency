pub mod activate_audio_async;
pub mod utils;

use core::slice;
use std::{
    mem, ptr,
    thread::{self, JoinHandle},
    time::Duration,
};
use windows_core::Interface;

use anyhow::Result;
use rtrb::{Consumer, Producer, RingBuffer, chunks::ChunkError};
use windows::Win32::{
    Media::{
        Audio::{
            AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
            AUDCLNT_STREAMFLAGS_LOOPBACK, AudioCategory_Media, AudioClientProperties,
            DEVICE_STATE_ACTIVE, EDataFlow, IAudioCaptureClient, IAudioClient, IAudioClient3,
            IAudioRenderClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, WAVEFORMATEX,
            eCapture, eRender,
        },
        Multimedia::WAVE_FORMAT_IEEE_FLOAT,
    },
    System::{
        Com::{
            CLSCTX_ALL, COINIT_MULTITHREADED, COINIT_SPEED_OVER_MEMORY, CoCreateInstance,
            CoInitializeEx,
        },
        Threading::{AvSetMmThreadCharacteristicsW, CreateEventW, WaitForSingleObject},
    },
};
use windows_strings::{HSTRING, w};

use crate::{
    activate_audio_async::capture_process_sync,
    utils::{IMMDeviceEx, WaveFormat, prompt},
};

// Spawn a COM multithreaded and set MMCSS Pro Audio task
pub fn spawn<F, T>(name: &str, f: F) -> JoinHandle<T>
where
    F: FnOnce() -> Result<T> + Send + 'static,
    T: Send + 'static,
{
    // let a = f().unwrap();
    // thread::spawn(|| a)
    thread::Builder::new()
        .name(name.into())
        .spawn(|| unsafe {
            CoInitializeEx(None, COINIT_SPEED_OVER_MEMORY | COINIT_MULTITHREADED)
                .ok()
                .unwrap();
            let mut task_idx = 0;
            AvSetMmThreadCharacteristicsW(w!("Pro Audio"), &mut task_idx).unwrap();
            println!("Registered for MMCSS Thread: TaskId = {task_idx}");
            f().unwrap()
        })
        .unwrap()
}

fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(None, COINIT_SPEED_OVER_MEMORY | COINIT_MULTITHREADED).ok()?;
        println!("Choose input type: ");
        println!("1: Device");
        println!("2: Process");
        let input = match prompt("Choice: ")? {
            1usize => {
                println!("Please select input device:");
                let input = prompt_device(eCapture)?;
                let input_id = input.GetId()?.to_string()?;
                Ok(input_id)
            }
            2usize => Err(prompt("Enter process id to capture: ")?),
            _ => panic!("Wrong choice!"),
        };

        println!("Please select output device:");
        let output = prompt_device(eRender)?;
        let ac: IAudioClient3 = output.Activate(CLSCTX_ALL, None)?;
        let wfx: WaveFormat = ac.GetMixFormat()?.into();

        let ac_capture = match input {
            Ok(input_id) => {
                let dev_enum: IMMDeviceEnumerator =
                    CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
                let dev = dev_enum.GetDevice(&HSTRING::from(input_id))?;
                let ac: IAudioClient3 = dev.Activate(CLSCTX_ALL, None)?;
                ac.cast()?
            }
            Err(pid) => {
                let ac = capture_process_sync(pid, true)?;
                ac
            }
        };

        let dev_enum: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let output_id = output.GetId()?.to_string()?;
        let dev = dev_enum.GetDevice(&HSTRING::from(output_id))?;
        let ac_render: IAudioClient3 = dev.Activate(CLSCTX_ALL, None)?;

        let mut ps = PipeStreamInfo::new(ac_capture, ac_render.cast()?, wfx)?;
        let mut task_idx = 0;
        AvSetMmThreadCharacteristicsW(w!("Pro Audio"), &mut task_idx).unwrap();
        println!("Registered for MMCSS Thread: TaskId = {task_idx}");
        ps.run()?;
        println!("Done");
        Ok(())
    }
}

pub struct PipeStreamInfo {
    capture: Producer<u8>,
    capture_client: IAudioClient,
    capture_info: InitInfo,
    render: Consumer<u8>,
    render_client: IAudioClient,
    render_info: InitInfo,
    ev: windows::Win32::Foundation::HANDLE,
    #[allow(unused)]
    wfx: WaveFormat,
}

impl PipeStreamInfo {
    pub fn new(capture: IAudioClient, render: IAudioClient, wfx: WaveFormat) -> Result<Self> {
        unsafe {
            let ev = CreateEventW(None, false, false, None)?;
            println!("Initialising input... ");
            let capture_info = init_ac(&capture, Some(wfx), ev)?;

            println!("Initialising output... ");
            let render_info = init_ac(&render, Some(wfx), ev)?;
            let (capture2, render2) = RingBuffer::new(480000 * 2);

            Ok(Self {
                capture_client: capture,
                capture_info,
                render_client: render,
                render_info,
                ev,
                wfx,
                capture: capture2,
                render: render2,
            })
        }
    }

    pub fn run(&mut self) -> Result<()> {
        unsafe {
            let cac: IAudioCaptureClient = self.capture_client.GetService()?;
            let crc: IAudioRenderClient = self.render_client.GetService()?;
            loop {
                WaitForSingleObject(self.ev, 2);
                loop {
                    while !self.render(&crc)? {}
                    if self.render.slots() > 0 {
                        break;
                    }
                    if self.capture(&cac)? {
                        break;
                    }
                }
            }
        }
    }

    // bool: Wait for signal
    fn capture(&mut self, cac: &IAudioCaptureClient) -> Result<bool> {
        unsafe {
            let mut cbuf = ptr::null_mut();
            let mut ftr = 0;
            let mut flags = 0;
            cac.GetBuffer(&mut cbuf, &mut ftr, &mut flags, None, None)?;
            if flags != 0 {
                println!("Capture flag not 0: {flags}");
            }
            if cbuf.is_null() {
                return Ok(true);
            }

            let rbuf = slice::from_raw_parts(cbuf, ftr as usize * self.capture_info.block as usize);
            match self.capture.write_chunk_uninit(rbuf.len()) {
                Ok(slot) => {
                    slot.fill_from_iter(rbuf.iter().copied());
                    cac.ReleaseBuffer(ftr)?;
                }
                Err(ChunkError::TooFewSlots(_)) => {
                    cac.ReleaseBuffer(0)?;
                }
            };

            let nps = cac.GetNextPacketSize()?;
            return Ok(nps == 0);
        }
    }

    // bool: Wait for signal
    fn render(&mut self, crc: &IAudioRenderClient) -> Result<bool> {
        unsafe {
            let padding = self.render_client.GetCurrentPadding()?;
            let available = self.render_info.buf_size - padding;
            if available == 0 {
                return Ok(true);
            }
            let cbuf = crc.GetBuffer(available)?;
            let rbuf = slice::from_raw_parts_mut(
                cbuf,
                available as usize * self.render_info.block as usize,
            );
            let slots = self.render.slots();
            let frames = slots * 1000
                / self.render_info.block as usize
                / (*self.render_info.wfx).nSamplesPerSec as usize;
            if frames > 30 {
                println!("warn: latency atm: {frames}ms");
            }
            let can_write = rbuf.len().min(slots);
            let slot = self.render.read_chunk(can_write)?;
            let data = slot.as_slices().0;
            rbuf[..data.len()].copy_from_slice(data);
            crc.ReleaseBuffer(data.len() as u32 / self.render_info.block, 0)?;
            slot.commit_all();
            return Ok(self.render.slots() == 0);
        }
    }
}

pub struct InitInfo {
    pub block: u32,
    pub wfx: WaveFormat,
    pub min_period: u32,
    pub buf_size: u32,
}

fn init_ac(
    ac: &IAudioClient,
    wfx: Option<WaveFormat>,
    ev: windows::Win32::Foundation::HANDLE,
) -> Result<InitInfo> {
    unsafe {
        let ac3: Option<IAudioClient3> = ac
            .cast()
            .inspect_err(|_| println!("This client does not support IAudioClient3!"))
            .ok();

        let wfx = wfx.unwrap_or(ac.GetMixFormat().map(|x| x.into()).unwrap_or_else(|_| {
            println!("This client doesnt support GetMixFormat");
            let wfx_new = WAVEFORMATEX {
                wFormatTag: WAVE_FORMAT_IEEE_FLOAT as u16,
                nChannels: 2,
                nSamplesPerSec: 48000,
                nAvgBytesPerSec: 384000,
                nBlockAlign: 8,
                wBitsPerSample: 32,
                cbSize: 22,
            };

            WaveFormat::Ex(wfx_new)
        }));
        println!("wave format: {:#?}", wfx);

        let min_period = if let Some(ac) = &ac3 {
            let mut props = AudioClientProperties::default();
            props.cbSize = mem::size_of_val(&props) as u32;
            props.eCategory = AudioCategory_Media;
            ac.SetClientProperties(&props)?;

            let mut default_period = 0;
            let mut fundamental_period = 0;
            let mut min_period = 0;
            let mut max_period = 0;

            ac.GetSharedModeEnginePeriod(
                wfx.as_mut_ptr(),
                &mut default_period,
                &mut fundamental_period,
                &mut min_period,
                &mut max_period,
            )?;

            let input_latency = (min_period as f64 * 1000f64) / (*wfx).nSamplesPerSec as f64;
            println!("default_period = {default_period}");
            println!("fundamental_period = {fundamental_period}");
            println!("min_period = {min_period}");
            println!("max_period = {max_period}");
            println!("latency = {input_latency}ms");
            ac.InitializeSharedAudioStream(
                AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                min_period,
                wfx.as_mut_ptr(),
                None,
            )?;
            min_period
        } else {
            println!("latency = 10ms");
            ac.Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                0,
                0,
                wfx.as_mut_ptr(),
                None,
            )?;
            10
        };

        let bfs = ac.GetBufferSize()?;
        println!("buffer size = {bfs}");

        ac.SetEventHandle(ev)?;
        ac.Start()?;

        Ok(InitInfo {
            block: (*wfx).nBlockAlign as u32,
            buf_size: bfs,
            min_period,
            wfx,
        })
    }
}

fn prompt_device(flow: EDataFlow) -> Result<IMMDevice> {
    let devs = get_devices(flow)?;
    for (i, dev) in devs.iter().enumerate() {
        let name = dev.display_name()?;
        println!("{i:<2} {name}");
    }
    let choice: usize = utils::prompt("Choice: ")?;
    Ok(devs.into_iter().skip(choice).next().unwrap())
}

fn get_devices(flow: EDataFlow) -> Result<Vec<IMMDevice>> {
    unsafe {
        let dev_enum: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let devs = dev_enum.EnumAudioEndpoints(flow, DEVICE_STATE_ACTIVE)?;
        let count = devs.GetCount()?;
        let s: Result<Vec<_>, _> = (0..count).map(|x| devs.Item(x)).collect();
        Ok(s?)
    }
}

pub fn to_reference_time(d: Duration) -> i64 {
    (d.as_nanos() / 100) as i64
}
