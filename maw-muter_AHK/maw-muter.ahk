; MAW-MUTER.ahk (Credits: VA.ahk() & mute_current_application())
global DeviceCount := 0
global DAEList := []

; to exclude devices create file mmae.txt and place keywords like this: Microphone,Speaker,Headphone
global excludeKeywords := []
LoadExcludeKeywords() {
    global excludeKeywords
    MMAEPATH1 := A_ScriptDir . "\mmae.txt"  ; Path to the text file
    MMAEPATH2 := A_ScriptDir . "\Config\mmae.txt"  ; Path to the text file

    if FileExist(MMAEPATH1) {
        FileRead, fileContent, %MMAEPATH1%
        excludeKeywords := StrSplit(fileContent, ",")
    }
    else if FileExist(MMAEPATH2)
    {
        FileRead, fileContent, %MMAEPATH2%
        excludeKeywords := StrSplit(fileContent, ",")
    }
}
LoadExcludeKeywords()

MAWAHK(ProcessName) {
    if !(Volume := GetVolumeObjectByName(ProcessName)) {
        return
    }
    
    VA_ISimpleAudioVolume_GetMute(Volume, Mute)  ; Get mute state
    VA_ISimpleAudioVolume_SetMute(Volume, !Mute) ; Toggle mute state
    ObjRelease(Volume)
    return
}

MAWAHKPID(PID) {
    if !(Volume := GetVolumeObjectByPID(PID)) {
        return
    }

    VA_ISimpleAudioVolume_GetMute(Volume, Mute)  ; Get mute state
    VA_ISimpleAudioVolume_SetMute(Volume, !Mute) ; Toggle mute state
    ObjRelease(Volume)
    return
}

GetDeviceCount() {

    global DeviceCount, DAEList
    DeviceCount := 0
    DAEList := []  ; Initialize an array to store device names
    
    Loop {
        ; Get the audio device
        DAE := VA_GetDevice(A_Index) ; Adjust index to start from 1

        ; If device is not found, exit loop
        if (!DAE) {
            ;MsgBox, Devices found: %DeviceCount%
            break
        }

        DeviceName := VA_GetDeviceName(DAE)

        ; Check if the device name contains any of the exclusion keywords
        shouldExclude := false
        for index, keyword in excludeKeywords {
            if (InStr(DeviceName, keyword) > 0) {
                ;MsgBox, Excluding: %DeviceName%
                shouldExclude := true
                break  ; Exit the loop if a match is found
            }
        }

        if (shouldExclude) {
            ; If it matches, skip this device
            ObjRelease(DAE)  ; Release the COM object if not used
            continue
        }

        ; If the device is not excluded, add to the list and increment the count
        DAEList.Push(A_Index)
        DeviceCount++

        ; Display the device name for verification (optional)
        ;MsgBox, %DeviceName% %DAE% %A_Index%

        ; Safety check to prevent endless loop
        if (DeviceCount > 100) {
            MsgBox, DeviceCount exceeded maximum device count (100). Please restart computer and try again. If the issue persists, please open an issue on GitHub's maw-muter repo.
            break
        }
        
        ; Release the COM object after use
        ObjRelease(DAE)
    }

    ; Return both DeviceCount and the list of devices
    return {Count: DeviceCount, Devices: DAEList}
}
GetDeviceCount()

GetVolumeObjectByName(targetExeName) {
    return GetVolumeObject(targetExeName, "name")
}

GetVolumeObjectByPID(targetPID) {
    return GetVolumeObject(targetPID, "pid")
}

GetVolumeObject(target, mode) {
    global DAEList
    static IID_IASM2 := "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}"
    , IID_IASC2 := "{bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}"
    , IID_ISAV := "{87CE5498-68D6-44E5-9215-6DA47EF883D8}"

    ; Get all audio devices
    Loop, % DAEList.MaxIndex()
    {
        DeviceNumber := DAEList[A_Index]  ; Get the device name from the list
        DAE := VA_GetDevice(DeviceNumber)

        if (DAE)
        {
            ; Check if the device is active and a rendering endpoint
            VA_IMMDevice_GetState(DAE, State)
            VA_IConnector_GetDataFlow(DAE, DataFlow)

            if (State == 1 && DataFlow == 0)  ; Check if the device is active and rendering
            {
                ; Activate the session manager
                VA_IMMDevice_Activate(DAE, IID_IASM2, 0, 0, IASM2)

                ; Enumerate sessions for the current device
                VA_IAudioSessionManager2_GetSessionEnumerator(IASM2, IASE)
                VA_IAudioSessionEnumerator_GetCount(IASE, Count)

                ; Search for an audio session with the required name for the current device
                Loop, % Count
                {
                    VA_IAudioSessionEnumerator_GetSession(IASE, A_Index-1, IASC)
                    IASC2 := ComObjQuery(IASC, IID_IASC2)

                    ; If IAudioSessionControl2 is queried successfully
                    if (IASC2)
                    {
                        VA_IAudioSessionControl2_GetProcessID(IASC2, SPID)

                        ; If the process name matches the one we are looking for
                        if ((mode == "name" && GetProcessNameFromPID(SPID) == target) || (mode == "pid" && SPID == target))
                        {
                            ; Check if the session is active before retrieving volume interface
                            VA_IAudioSessionControl_GetState(IASC2, SessionState)
                            if (SessionState == 1) ; AudioSessionStateActive
                            {
                                ISAV := ComObjQuery(IASC2, IID_ISAV)
                                if (ISAV)
                                {
                                    ObjRelease(IASC2)
                                    ObjRelease(IASC)
                                    ObjRelease(IASE)
                                    ObjRelease(IASM2)
                                    ObjRelease(DAE)
                                    return ISAV
                                    break
                                }
                            }
                        }
                        ObjRelease(IASC2)
                    }
                    ObjRelease(IASC)
                }
                ObjRelease(IASE)
            }
            ObjRelease(IASM2)
            ObjRelease(DAE)
        }
    }
    ; If no active audio session is found for the PID, get the process name and retry
    if (mode == "pid")
    {
        processName := GetProcessNameFromPID(target)
        ; MsgBox, % "No active audio session found for PID: " target "`nRetrying with process name: " processName
        return GetVolumeObject(processName, "name")
    }
    ; MsgBox, No active audio session found for the specified process: %targetExeName%
    GetDeviceCount()
    return ; Return 0 if there's an issue retrieving the interface
}

GetProcessNameFromPID(PID)
{
    hProcess := DllCall("OpenProcess", "UInt", 0x0400 | 0x0010, "Int", false, "UInt", PID)
    VarSetCapacity(ExeName, 260, 0)
    DllCall("Psapi.dll\GetModuleFileNameEx", "UInt", hProcess, "UInt", 0, "Str", ExeName, "UInt", 260)
    DllCall("CloseHandle", "UInt", hProcess)
    return SubStr(ExeName, InStr(ExeName, "\", false, -1) + 1)
}

;VA.ahk (stripped down version by tfurci)

VA_GetDevice(device_desc="playback")
{
    if ( r:= DllCall("ole32\CoCreateInstance"
                , "ptr", VA_GUID(CLSID_MMDeviceEnumerator, "{BCDE0395-E52F-467C-8E3D-C4579291692E}")
                , "ptr", 0, "uint", 21
                , "ptr", VA_GUID(IID_IMMDeviceEnumerator, "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
                , "ptr*", deviceEnumerator)) != 0
        return 0
    
    device := 0
    
    ; deviceEnumerator->GetDevice(device_id, [out] device)
    if DllCall(NumGet(NumGet(deviceEnumerator+0)+5*A_PtrSize), "ptr", deviceEnumerator, "wstr", device_desc, "ptr*", device) = 0
        goto VA_GetDevice_Return
    
    if device_desc is integer
    {
        m2 := device_desc
        if m2 >= 4096 ; Probably a device pointer, passed here indirectly via VA_GetAudioMeter or such.
            return m2, ObjAddRef(m2)
    }
    else
        RegExMatch(device_desc, "(.*?)\s*(?::(\d+))?$", m)
    
    if m1 in playback,p
        m1 := "", flow := 0 ; eRender
    else if m1 in capture,c
        m1 := "", flow := 1 ; eCapture
    else if (m1 . m2) = ""  ; no name or number specified
        m1 := "", flow := 0 ; eRender (default)
    else
        flow := 2 ; eAll
    
    if (m1 . m2) = ""   ; no name or number (maybe "playback" or "capture")
    {   ; deviceEnumerator->GetDefaultAudioEndpoint(dataFlow, role, [out] device)
        DllCall(NumGet(NumGet(deviceEnumerator+0)+4*A_PtrSize), "ptr",deviceEnumerator, "uint",flow, "uint",0, "ptr*",device)
        goto VA_GetDevice_Return
    }

    ; deviceEnumerator->EnumAudioEndpoints(dataFlow, stateMask, [out] devices)
    DllCall(NumGet(NumGet(deviceEnumerator+0)+3*A_PtrSize), "ptr",deviceEnumerator, "uint",flow, "uint",1, "ptr*",devices)
    
    ; devices->GetCount([out] count)
    DllCall(NumGet(NumGet(devices+0)+3*A_PtrSize), "ptr",devices, "uint*",count)
    
    if m1 =
    {   ; devices->Item(m2-1, [out] device)
        DllCall(NumGet(NumGet(devices+0)+4*A_PtrSize), "ptr",devices, "uint",m2-1, "ptr*",device)
        goto VA_GetDevice_Return
    }
    
    index := 0
    Loop % count
        ; devices->Item(A_Index-1, [out] device)
        if DllCall(NumGet(NumGet(devices+0)+4*A_PtrSize), "ptr",devices, "uint",A_Index-1, "ptr*",device) = 0
            if InStr(VA_GetDeviceName(device), m1) && (m2 = "" || ++index = m2)
                goto VA_GetDevice_Return
            else
                ObjRelease(device), device:=0

VA_GetDevice_Return:
    ObjRelease(deviceEnumerator)
    if devices
        ObjRelease(devices)
    
    return device ; may be 0
}

VA_GetDeviceName(device)
{
    static PKEY_Device_FriendlyName
    if !VarSetCapacity(PKEY_Device_FriendlyName)
        VarSetCapacity(PKEY_Device_FriendlyName, 20)
        ,VA_GUID(PKEY_Device_FriendlyName :="{A45C254E-DF1C-4EFD-8020-67D146A850E0}")
        ,NumPut(14, PKEY_Device_FriendlyName, 16)
    VarSetCapacity(prop, 16)
    VA_IMMDevice_OpenPropertyStore(device, 0, store)
    ; store->GetValue(.., [out] prop)
    DllCall(NumGet(NumGet(store+0)+5*A_PtrSize), "ptr", store, "ptr", &PKEY_Device_FriendlyName, "ptr", &prop)
    ObjRelease(store)
    VA_WStrOut(deviceName := NumGet(prop,8))
    return deviceName
}

; Convert string to binary GUID structure.
VA_GUID(ByRef guid_out, guid_in="%guid_out%") {
    if (guid_in == "%guid_out%")
        guid_in :=   guid_out
    if  guid_in is integer
        return guid_in
    VarSetCapacity(guid_out, 16, 0)
	DllCall("ole32\CLSIDFromString", "wstr", guid_in, "ptr", &guid_out)
	return &guid_out
}

VA_WStrOut(ByRef str) {
    str := StrGet(ptr := str, "UTF-16")
    DllCall("ole32\CoTaskMemFree", "ptr", ptr)  ; FREES THE STRING.
}

VA_IMMDevice_OpenPropertyStore(this, Access, ByRef Properties) {
    return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint", Access, "ptr*", Properties)
}

VA_IMMDevice_GetState(this, ByRef State) {
    return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "uint*", State)
}

VA_IConnector_GetDataFlow(this, ByRef Flow) {
    return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "int*", Flow)
}

VA_IMMDevice_Activate(this, iid, ClsCtx, ActivationParams, ByRef Interface) {
    return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "ptr", VA_GUID(iid), "uint", ClsCtx, "uint", ActivationParams, "ptr*", Interface)
}

VA_IAudioSessionManager2_GetSessionEnumerator(this, ByRef SessionEnum) {
    return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "ptr*", SessionEnum)
}

VA_IAudioSessionEnumerator_GetCount(this, ByRef SessionCount) {
    return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int*", SessionCount)
}

VA_IAudioSessionEnumerator_GetSession(this, SessionCount, ByRef Session) {
    return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "int", SessionCount, "ptr*", Session)
}

VA_IAudioSessionControl2_GetProcessId(this, ByRef pid) {
    return DllCall(NumGet(NumGet(this+0)+14*A_PtrSize), "ptr", this, "uint*", pid)
}

VA_IAudioSessionControl_GetState(this, ByRef State) {
    return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int*", State)
}

VA_ISimpleAudioVolume_SetMasterVolume(this, ByRef fLevel, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "float", fLevel, "ptr", VA_GUID(GuidEventContext))
}

VA_ISimpleAudioVolume_GetMasterVolume(this, ByRef fLevel) {
    return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "float*", fLevel)
}

VA_ISimpleAudioVolume_SetMute(this, ByRef Muted, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "int", Muted, "ptr", VA_GUID(GuidEventContext))
}

VA_ISimpleAudioVolume_GetMute(this, ByRef Muted) {
    return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "int*", Muted)
}
