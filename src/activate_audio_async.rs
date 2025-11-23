use std::sync::Mutex;

use future_handles::{
    self, HandleError,
    sync::{self, CompleteHandle},
};
use thiserror::Error;
use windows::{
    Win32::{
        Media::Audio::{
            ActivateAudioInterfaceAsync, IActivateAudioInterfaceAsyncOperation,
            IActivateAudioInterfaceCompletionHandler,
            IActivateAudioInterfaceCompletionHandler_Impl,
        },
        System::Com::StructuredStorage::PROPVARIANT,
    },
    core::{HRESULT, IUnknown, Interface, implement},
};

pub async unsafe fn activate_audio_interface_async<P0, Out>(
    deviceinterfacepath: P0,
    activationparams: ::core::option::Option<*const PROPVARIANT>,
) -> Result<Out, ActivationError>
where
    P0: ::windows::core::Param<::windows::core::PCWSTR>,
    Out: Interface,
{
    let (future, handle) = sync::create();

    let completionhandler: IActivateAudioInterfaceCompletionHandler =
        AsyncHandler(Mutex::new(Some(handle))).into();
    let result = ActivateAudioInterfaceAsync(
        deviceinterfacepath,
        &Out::IID,
        activationparams,
        &completionhandler,
    )?;
    future.await?;

    let mut hr = HRESULT(0);
    // let mut hr: MaybeUninit<HRESULT> = MaybeUninit::uninit();
    let mut ai: Option<IUnknown> = None;
    result.GetActivateResult(&mut hr, &mut ai)?;

    if let Some(comi) = ai {
        Ok(comi.cast()?)
    } else {
        let err = windows::core::Error::from(hr);
        Err(err.into())
    }
}

#[implement(IActivateAudioInterfaceCompletionHandler)]
struct AsyncHandler(Mutex<Option<CompleteHandle<()>>>);

impl IActivateAudioInterfaceCompletionHandler_Impl for AsyncHandler_Impl {
    fn ActivateCompleted(
        &self,
        _: windows::core::Ref<IActivateAudioInterfaceAsyncOperation>,
    ) -> windows::core::Result<()> {
        let val = self.0.lock().unwrap().take();
        if let Some(handle) = val {
            handle.complete(());
        }
        Ok(())
    }
}

#[derive(Debug, Error)]
pub enum ActivationError {
    #[error(transparent)]
    WindowError(windows::core::Error),

    #[error(transparent)]
    HandleError(HandleError),
}

impl From<HandleError> for ActivationError {
    fn from(value: HandleError) -> Self {
        ActivationError::HandleError(value)
    }
}

impl From<windows::core::Error> for ActivationError {
    fn from(value: windows::core::Error) -> Self {
        ActivationError::WindowError(value)
    }
}
