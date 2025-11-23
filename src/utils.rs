use std::{
    fmt::{Debug, Display},
    fs,
    io::{self, BufRead, BufReader, Write},
    mem,
    ops::Deref,
    str::FromStr,
    sync::{LazyLock, Mutex},
};

use anyhow::{Result, anyhow};
use extension_trait::extension_trait;
use windows::Win32::{
    Devices::FunctionDiscovery::PKEY_Device_FriendlyName,
    Media::{
        Audio::{IMMDevice, WAVEFORMATEX, WAVEFORMATEXTENSIBLE},
        KernelStreaming::WAVE_FORMAT_EXTENSIBLE,
    },
    System::Com::{CoTaskMemFree, STGM_READWRITE},
};

#[extension_trait]
pub impl IMMDeviceEx for IMMDevice {
    fn display_name(&self) -> Result<impl Display> {
        unsafe {
            let props = self.OpenPropertyStore(STGM_READWRITE)?;
            let name = props.GetValue(&PKEY_Device_FriendlyName)?;
            Ok(name)
        }
    }
}

pub fn prompt_with<T: FromStr>(q: impl Display, input: &mut impl BufRead) -> Result<T>
where
    T::Err: std::error::Error + Send + Sync + 'static,
{
    print!("{q}");
    io::stdout().flush()?;
    let mut buf = String::new();
    input.read_line(&mut buf)?;
    Ok(buf.trim().parse::<T>()?)
}

pub fn prompt_stdio<T: FromStr>(q: impl Display) -> Result<T>
where
    T::Err: std::error::Error + Send + Sync + 'static,
{
    prompt_with(q, &mut std::io::stdin().lock())
}

pub fn prompt<T: FromStr>(q: impl Display) -> Result<T>
where
    T::Err: std::error::Error + Send + Sync + 'static,
{
    if let Some(lp) = LP.as_ref() {
        let mut guard = lp.lock().map_err(|_| anyhow!("cannot lock"))?;
        prompt_with(q, &mut &mut *guard)
    } else {
        prompt_with(q, &mut std::io::stdin().lock())
    }
}

static LP: LazyLock<Option<Box<Mutex<dyn BufRead + Send + Sync>>>> = LazyLock::new(|| {
    fs::File::open("./stdio.txt")
        .ok()
        .map(|x| BufReader::new(x))
        .map(|x| Mutex::new(x))
        .map(|x| Box::new(x) as Box<Mutex<dyn BufRead + Send + Sync>>)
});

#[extension_trait]
pub impl Wftex for WAVEFORMATEX {
    fn debug(&self) -> impl Debug + '_ {
        if self.wFormatTag == WAVE_FORMAT_EXTENSIBLE as u16 {
            unsafe {
                let wrapper = DbgWrapper(&*(self as *const _ as *const WAVEFORMATEXTENSIBLE));
                return Box::new(wrapper) as Box<dyn Debug>;
            }
        } else {
            let wrapper = DbgWrapper(self);
            return Box::new(wrapper) as Box<dyn Debug>;
        }
    }

    fn debug_no_downcast(&self) -> impl Debug + '_ {
        DbgWrapper(self)
    }
}

#[derive(Clone, Copy)]
pub enum WaveFormat {
    Ex(WAVEFORMATEX),
    Extensible(WAVEFORMATEXTENSIBLE),
}

impl Deref for WaveFormat {
    type Target = WAVEFORMATEX;

    fn deref(&self) -> &Self::Target {
        unsafe { &*(self.as_mut_ptr() as *const _) }
    }
}

impl Debug for WaveFormat {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let ptr = self.as_mut_ptr();
        unsafe { (*ptr).debug().fmt(f) }
    }
}

impl WaveFormat {
    pub fn as_mut_ptr(&self) -> *mut WAVEFORMATEX {
        match self {
            WaveFormat::Ex(waveformatex) => waveformatex as *const _ as *mut _,
            WaveFormat::Extensible(waveformatextensible) => {
                waveformatextensible as *const _ as *mut WAVEFORMATEX
            }
        }
    }
}

/// This will move into an owned type and call CoTaskMemFree on the pointer
impl From<*mut WAVEFORMATEX> for WaveFormat {
    fn from(value: *mut WAVEFORMATEX) -> Self {
        unsafe {
            let rs = if (*value).wFormatTag == WAVE_FORMAT_EXTENSIBLE as u16 {
                Self::Extensible(*(value as *mut WAVEFORMATEXTENSIBLE))
            } else {
                Self::Ex(*value)
            };

            CoTaskMemFree(Some(value as *mut _));
            rs
        }
    }
}

#[extension_trait]
pub impl Wft2ex for WAVEFORMATEXTENSIBLE {
    fn debug(&self) -> impl Debug {
        DbgWrapper(self)
    }
}

pub struct DbgWrapper<'a, T>(&'a T);

impl<'a> Debug for DbgWrapper<'a, WAVEFORMATEX> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WAVEFORMATEX")
            .field("wFormatTag", &{ self.0.wFormatTag })
            .field("nChannels", &{ self.0.nChannels })
            .field("nSamplesPerSec", &{ self.0.nSamplesPerSec })
            .field("nAvgBytesPerSec", &{ self.0.nAvgBytesPerSec })
            .field("nBlockAlign", &{ self.0.nBlockAlign })
            .field("wBitsPerSample", &{ self.0.wBitsPerSample })
            .field("cbSize", &{ self.0.cbSize })
            .finish()
    }
}

impl<'a> Debug for DbgWrapper<'a, WAVEFORMATEXTENSIBLE> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        unsafe {
            f.debug_struct("WAVEFORMATEXTENSIBLE")
                .field("Format", &{ self.0.Format.debug_no_downcast() })
                .field("Samples", &mem::transmute::<_, u16>(self.0.Samples))
                .field("dwChannelMask", &{ self.0.dwChannelMask })
                .field("SubFormat", &{ self.0.SubFormat })
                .finish()
        }
    }
}
