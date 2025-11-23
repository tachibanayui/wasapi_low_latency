pub mod activate_audio_async;
pub mod utils;

use core::slice;
use std::{
    mem,
    pin::pin,
    ptr,
    task::Context,
    thread::{self, JoinHandle},
    time::Duration,
};
use windows_core::Interface;

use anyhow::Result;
use rtrb::{RingBuffer, chunks::ChunkError};
use tokio::runtime::Runtime;
use windows::Win32::{
    Media::{
        Audio::{
            AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
            AUDCLNT_STREAMFLAGS_LOOPBACK, AUDIOCLIENT_ACTIVATION_PARAMS,
            AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK, AudioCategory_Media,
            AudioClientProperties, DEVICE_STATE_ACTIVE, EDataFlow, IAudioCaptureClient,
            IAudioClient, IAudioClient3, IAudioRenderClient, IMMDevice, IMMDeviceEnumerator,
            MMDeviceEnumerator, PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE,
            VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, WAVEFORMATEX, WAVEFORMATEXTENSIBLE, eCapture,
            eRender,
        },
        Multimedia::WAVE_FORMAT_IEEE_FLOAT,
    },
    System::{
        Com::{
            CLSCTX_ALL, COINIT_MULTITHREADED, COINIT_SPEED_OVER_MEMORY, CoCreateInstance,
            CoInitializeEx, CoTaskMemAlloc, StructuredStorage::PROPVARIANT,
        },
        Threading::{AvSetMmThreadCharacteristicsW, CreateEventW, INFINITE, WaitForSingleObject},
        Variant::VT_BLOB,
    },
};
use windows_strings::{HSTRING, w};

use crate::{
    activate_audio_async::activate_audio_interface_async,
    utils::{IMMDeviceEx, WaveFormat, Wftex},
};

fn spawn<F, T>(name: &str, f: F) -> JoinHandle<T>
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
        let input = match utils::prompt("Choice: ")? {
            1usize => {
                println!("Please select input device:");
                let input = prompt_device(eCapture)?;
                let input_id = input.GetId()?.to_string()?;
                Ok(input_id)
            }
            2usize => Err(utils::prompt("Enter process id to capture: ")?),
            // I'm too lazy to handle here properly ^^
            _ => panic!("Wrong choice!"),
        };

        println!("Please select output device:");
        let output = prompt_device(eRender)?;
        let ac: IAudioClient3 = output.Activate(CLSCTX_ALL, None)?;
        let mut wfx: WaveFormat = ac.GetMixFormat()?.into();

        let (mut capture, mut render) = RingBuffer::new(48000 * 2);
        let capture_thd = spawn("capture", move || {
            let ac = match input {
                Ok(input_id) => {
                    let dev_enum: IMMDeviceEnumerator =
                        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
                    let dev = dev_enum.GetDevice(&HSTRING::from(input_id))?;
                    let ac: IAudioClient3 = dev.Activate(CLSCTX_ALL, None)?;
                    ac.cast()?
                }
                Err(pid) => {
                    let rt = Runtime::new().unwrap();
                    let ac = rt.block_on(process_loopback(pid))?;
                    ac
                }
            };

            println!("Initialising input... ");
            let info = init_ac(&ac, Some(wfx.as_mut_ptr()))?;
            let cac: IAudioCaptureClient = ac.GetService()?;
            'main: loop {
                WaitForSingleObject(info.ev, INFINITE);

                loop {
                    let mut cbuf = ptr::null_mut();
                    let mut ftr = 0;
                    let mut flags = 0;
                    cac.GetBuffer(&mut cbuf, &mut ftr, &mut flags, None, None)?;
                    if flags != 0 {
                        println!("Capture flag not 0: {flags}");
                    }
                    if cbuf.is_null() {
                        continue;
                    }
                    let rbuf = slice::from_raw_parts(cbuf, ftr as usize * info.block as usize);
                    let slot = match capture.write_chunk_uninit(rbuf.len()) {
                        Ok(ok) => ok,
                        Err(ChunkError::TooFewSlots(_)) => {
                            cac.ReleaseBuffer(0)?;
                            continue;
                        }
                    };
                    slot.fill_from_iter(rbuf.iter().copied());
                    cac.ReleaseBuffer(ftr)?;

                    let nps = cac.GetNextPacketSize()?;
                    if nps == 0 {
                        continue 'main;
                    }
                }
            }

            #[allow(unreachable_code)]
            Ok::<_, anyhow::Error>(())
        });

        let output_id = output.GetId()?.to_string()?;
        let render_thd = spawn("render", move || {
            let dev_enum: IMMDeviceEnumerator =
                CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

            let dev = dev_enum.GetDevice(&HSTRING::from(output_id))?;
            let ac: IAudioClient3 = dev.Activate(CLSCTX_ALL, None)?;
            println!("Initialising output... ");
            let info = init_ac(&ac, None)?;
            let cac: IAudioRenderClient = ac.GetService()?;
            'main: loop {
                WaitForSingleObject(info.ev, INFINITE);
                loop {
                    let padding = ac.GetCurrentPadding()?;
                    let available = info.buf_size - padding;
                    if available == 0 {
                        continue 'main;
                    }
                    let cbuf = cac.GetBuffer(available)?;
                    let rbuf =
                        slice::from_raw_parts_mut(cbuf, available as usize * info.block as usize);
                    let slots = render.slots();
                    let frames =
                        slots * 1000 / info.block as usize / (*info.wfx).nSamplesPerSec as usize;
                    println!("latency atm: {frames}ms");
                    let can_write = rbuf.len().min(slots);
                    let slot = render.read_chunk(can_write)?;
                    let data = slot.as_slices().0;
                    rbuf[..data.len()].copy_from_slice(data);
                    cac.ReleaseBuffer(data.len() as u32 / info.block, 0)?;
                    slot.commit_all();
                }
            }

            #[allow(unreachable_code)]
            return Ok::<_, anyhow::Error>(());
        });

        dbg!(render_thd.join()).unwrap();
        dbg!(capture_thd.join()).unwrap();
        println!("Done");
        Ok(())
    }
}

struct InitInfo {
    block: u32,
    wfx: *mut WAVEFORMATEX,
    min_period: u32,
    ev: windows::Win32::Foundation::HANDLE,
    buf_size: u32,
}

fn init_ac(ac: &IAudioClient, mfx: Option<*mut WAVEFORMATEX>) -> Result<InitInfo> {
    unsafe {
        let ac3: Option<IAudioClient3> = ac
            .cast()
            .inspect_err(|_| println!("This client does not support IAudioClient3!"))
            .ok();

        let wfx = mfx.unwrap_or(ac.GetMixFormat().unwrap_or_else(|_| {
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

            let wfxptr = CoTaskMemAlloc(mem::size_of_val(&wfx_new)) as *mut WAVEFORMATEX;
            *wfxptr = wfx_new;
            wfxptr
        }));
        println!("wave format: {:#?}", (*wfx).debug());

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
                wfx,
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
                wfx,
                None,
            )?;
            min_period
        } else {
            println!("latency = 10ms");
            ac.Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                to_reference_time(Duration::from_millis(2)),
                0,
                wfx,
                None,
            )?;
            10
        };

        let bfs = ac.GetBufferSize()?;
        println!("buffer size = {bfs}");

        let ev = CreateEventW(None, false, false, None)?;
        ac.SetEventHandle(ev)?;
        ac.Start()?;

        Ok(InitInfo {
            block: (*wfx).nBlockAlign as u32,
            buf_size: bfs,
            ev: ev,
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

/// `PROPVARIANT` referencing `AUDIOCLIENT_ACTIVATION_PARAMS`! Do not drop it before the returned value
/// Treat it like into_propvariant(client: &AUDIOCLIENT_ACTIVATION_PARAMS) -> PROPVARIANT + '_
unsafe fn into_propvariant(client: &AUDIOCLIENT_ACTIVATION_PARAMS) -> PROPVARIANT {
    use std::{mem, ptr};

    let mut p = PROPVARIANT::default();
    // let vt = &p.Anonymous.Anonymous.vt as *const _ as *mut _;
    // let blob_size = &p.Anonymous.Anonymous.Anonymous.blob.cbSize as *const _ as *mut _;
    // let blob_data = &p.Anonymous.Anonymous.Anonymous.blob.pBlobData as *const _ as *mut AUDIOCLIENT_ACTIVATION_PARAMS;

    unsafe {
        (*p.Anonymous.Anonymous).vt = VT_BLOB;
        (*p.Anonymous.Anonymous).Anonymous.blob.cbSize =
            mem::size_of::<AUDIOCLIENT_ACTIVATION_PARAMS>() as u32;
        (*p.Anonymous.Anonymous).Anonymous.blob.pBlobData = client as *const _ as *mut u8;
        p
    }
}

async fn process_loopback(proc_no: u32) -> Result<IAudioClient> {
    let mut params = AUDIOCLIENT_ACTIVATION_PARAMS::default();
    params.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    params.Anonymous.ProcessLoopbackParams.ProcessLoopbackMode =
        PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE;
    params.Anonymous.ProcessLoopbackParams.TargetProcessId = proc_no;
    unsafe {
        let pv = into_propvariant(&params);
        let aud_client: IAudioClient =
            activate_audio_interface_async(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, Some(&pv))
                .await
                .unwrap();
        mem::forget(pv);
        Ok(aud_client)
    }
}

pub fn to_reference_time(d: Duration) -> i64 {
    (d.as_nanos() / 100) as i64
}
