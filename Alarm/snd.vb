Imports System
Imports System.Runtime.InteropServices

Public Class Sound

#Region "   Events"
    Public Event MuteChanged()
    Public Event VolumeChanged()
#End Region

#Region "   Constants"
    Private Const MMSYSERR_NOERROR As Integer = 0
    Private Const MAXPNAMELEN As Integer = 32
    Private Const MIXER_LONG_NAME_CHARS As Integer = 64
    Private Const MIXER_SHORT_NAME_CHARS As Integer = 16
    Private Const MIXER_GETLINEINFOF_COMPONENTTYPE As Integer = &H3
    Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE As Integer = &H2
    Private Const MIXER_GETCONTROLDETAILSF_VALUE As Integer = &H0
    Private Const MIXER_SETCONTROLDETAILSF_VALUE As Integer = &H0
    Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST As Integer = &H0
    Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS As Integer = MIXERLINE_COMPONENTTYPE_DST_FIRST + 4
    'Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST As Integer = &H1000
    'Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE As Integer = MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3
    'Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE As Integer = MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2
    'Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
    Private Const MIXERCONTROL_CT_CLASS_FADER As Integer = &H50000000
    Private Const MIXERCONTROL_CT_UNITS_UNSIGNED As Integer = &H30000
    Private Const MIXERCONTROL_CT_CLASS_SWITCH As Integer = &H20000000
    Private Const MIXERCONTROL_CT_UNITS_BOOLEAN As Integer = &H10000
    Private Const MIXERCONTROL_CONTROLTYPE_BASS As Integer = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
    Private Const MIXERCONTROL_CONTROLTYPE_FADER As Integer = MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED
    Private Const MIXERCONTROL_CONTROLTYPE_VOLUME As Integer = MIXERCONTROL_CONTROLTYPE_FADER + 1
    Private Const MIXERCONTROL_CONTROLTYPE_MUTE As Integer = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
    Private Const MIXERCONTROL_CONTROLTYPE_TREBLE As Integer = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
    Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER As Integer = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
    Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN As Integer = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_UNITS_BOOLEAN)
#End Region

#Region "   API Calls"
    Private Declare Ansi Function mixerClose Lib "winmm.dll" (ByVal hmx As Integer) As Integer
    Private Declare Ansi Function mixerGetControlDetailsA Lib "winmm.dll" (ByVal hmxobj As Integer, ByRef pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Integer) As Integer
    Private Declare Ansi Function mixerGetDevCapsA Lib "winmm.dll" (ByVal uMxId As Integer, ByVal pmxcaps As MIXERCAPS, ByVal cbmxcaps As Integer) As Integer
    Private Declare Ansi Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Integer, ByVal pumxID As Integer, ByVal fdwId As Integer) As Integer
    Private Declare Ansi Function mixerGetLineControlsA Lib "winmm.dll" (ByVal hmxobj As Integer, ByRef pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Integer) As Integer
    Private Declare Ansi Function mixerGetLineInfoA Lib "winmm.dll" (ByVal hmxobj As Integer, ByRef pmxl As MIXERLINE, ByVal fdwInfo As Integer) As Integer
    Private Declare Ansi Function mixerGetNumDevs Lib "winmm.dll" () As Integer
    Private Declare Ansi Function mixerMessage Lib "winmm.dll" (ByVal hmx As Integer, ByVal uMsg As Integer, ByVal dwParam1 As Integer, ByVal dwParam2 As Integer) As Integer
    Private Declare Ansi Function mixerOpen Lib "winmm.dll" (ByRef phmx As Integer, ByVal uMxId As Integer, ByVal dwCallback As Integer, ByVal dwInstance As Integer, ByVal fdwOpen As Integer) As Integer
    Private Declare Ansi Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Integer, ByRef pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Integer) As Integer
    Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (<MarshalAs(UnmanagedType.I4)> ByVal hmxobj As Integer, ByRef pmxl As MIXERLINE, ByVal fdwInfo As Integer) As Integer
    Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (<MarshalAs(UnmanagedType.I4)> ByVal hmxobj As Integer, ByRef pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Integer) As Integer
#End Region

#Region "   Structures"
    Private Structure MIXERCAPS
        Public wMid As Integer
        Public wPid As Integer
        Public vDriverVersion As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAXPNAMELEN)> Public szPname As String
        Public fdwSupport As Integer
        Public cDestinations As Integer
    End Structure 'MIXERCAPS
       _
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERCONTROL
        <FieldOffset(0)> Public cbStruct As Integer           '  size in Byte of MIXERCONTROL
        <FieldOffset(4)> Public dwControlID As Integer        '  unique control id for mixer device
        <FieldOffset(8)> Public dwControlType As Integer      '  MIXERCONTROL_CONTROLTYPE_xxx
        <FieldOffset(12)> Public fdwControl As Integer         '  MIXERCONTROL_CONTROLF_xxx
        <FieldOffset(16)> Public cMultipleItems As Integer     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
        <FieldOffset(20), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MIXER_SHORT_NAME_CHARS)> Public szShortName As String ' * MIXER_SHORT_NAME_CHARS  ' short name of control
        <FieldOffset(36), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MIXER_LONG_NAME_CHARS)> Public szName As String '  * MIXER_LONG_NAME_CHARS ' Integer name of control
        <FieldOffset(100)> Public lMinimum As Integer           '  Minimum value
        <FieldOffset(104)> Public lMaximum As Integer           '  Maximum value
        <FieldOffset(108), MarshalAs(UnmanagedType.ByValArray, SizeConst:=11, ArraySubType:=UnmanagedType.AsAny)> Public reserved() As Integer      '  reserved structure space
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERCONTROLDETAILS
        <FieldOffset(0)> Public cbStruct As Integer       '  size in Byte of MIXERCONTROLDETAILS
        <FieldOffset(4)> Public dwControlID As Integer    '  control id to get/set details on
        <FieldOffset(8)> Public cChannels As Integer      '  number of channels in paDetails array
        <FieldOffset(12)> Public item As Integer           '  hwndOwner or cMultipleItems
        <FieldOffset(16)> Public cbDetails As Integer      '  size of _one_ details_XX struct
        <FieldOffset(20)> Public paDetails As IntPtr       '  pointer to array of details_XX structs
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERCONTROLDETAILS_UNSIGNED
        <FieldOffset(0)> Public dwValue As Integer        '  value of the control
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERLINE
        <FieldOffset(0)> Public cbStruct As Integer                '  size of MIXERLINE structure
        <FieldOffset(4)> Public dwDestination As Integer          '  zero based destination index
        <FieldOffset(8)> Public dwSource As Integer               '  zero based source index (if source)
        <FieldOffset(12)> Public dwLineID As Integer               '  unique line id for mixer device
        <FieldOffset(16)> Public fdwLine As Integer                '  state/information about line
        <FieldOffset(20)> Public dwUser As Integer                 '  driver specific information
        <FieldOffset(24)> Public dwComponentType As Integer        '  component type line connects to
        <FieldOffset(28)> Public cChannels As Integer              '  number of channels line supports
        <FieldOffset(32)> Public cConnections As Integer           '  number of connections (possible)
        <FieldOffset(36)> Public cControls As Integer              '  number of controls at this line
        <FieldOffset(40), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MIXER_SHORT_NAME_CHARS)> Public szShortName As String  ' * MIXER_SHORT_NAME_CHARS
        <FieldOffset(56), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MIXER_LONG_NAME_CHARS)> Public szName As String ' * MIXER_LONG_NAME_CHARS
        <FieldOffset(120)> Public dwType As Integer
        <FieldOffset(124)> Public dwDeviceID As Integer
        <FieldOffset(128)> Public wMid As Integer
        <FieldOffset(132)> Public wPid As Integer
        <FieldOffset(136)> Public vDriverVersion As Integer
        <FieldOffset(168), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MAXPNAMELEN)> Public szPname As String ' * MAXPNAMELEN
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERLINECONTROLS
        <FieldOffset(0)> Public cbStruct As Integer       '  size in Byte of MIXERLINECONTROLS
        <FieldOffset(4)> Public dwLineID As Integer       '  line id (from MIXERLINE.dwLineID)
        <FieldOffset(8)> Public dwControl As Integer      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
        <FieldOffset(12)> Public cControls As Integer      '  count of controls pmxctrl points to
        <FieldOffset(16)> Public cbmxctrl As Integer       '  size in Byte of _one_ MIXERCONTROL
        <FieldOffset(20)> Public pamxctrl As IntPtr       '  pointer to first MIXERCONTROL array
    End Structure

#End Region

    Private Shared Function GetVolumeControl(ByVal hmixer As Integer, ByVal componentType As Integer, ByVal ctrlType As Integer, ByRef mxc As MIXERCONTROL, ByRef vCurrentVol As Integer) As Boolean
        Dim mxlc As New MIXERLINECONTROLS
        Dim mxl As New MIXERLINE
        Dim pmxcd As New MIXERCONTROLDETAILS
        Dim du As New MIXERCONTROLDETAILS_UNSIGNED
        mxc = New MIXERCONTROL
        Dim rc As Integer
        Dim retValue As Boolean
        vCurrentVol = -1

        mxl.cbStruct = Marshal.SizeOf(mxl)
        mxl.dwComponentType = componentType

        rc = mixerGetLineInfoA(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)

        If MMSYSERR_NOERROR = rc Then
            Dim sizeofMIXERCONTROL As Integer = 152
            Dim ctrl As Integer = Marshal.SizeOf(GetType(MIXERCONTROL))
            mxlc.pamxctrl = Marshal.AllocCoTaskMem(sizeofMIXERCONTROL)
            mxlc.cbStruct = Marshal.SizeOf(mxlc)
            mxlc.dwLineID = mxl.dwLineID
            mxlc.dwControl = ctrlType
            mxlc.cControls = 1
            mxlc.cbmxctrl = sizeofMIXERCONTROL

            ' Allocate a buffer for the control 
            mxc.cbStruct = sizeofMIXERCONTROL

            ' Get the control 
            rc = mixerGetLineControlsA(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)

            If MMSYSERR_NOERROR = rc Then
                retValue = True

                ' Copy the control into the destination structure 
                mxc = CType(Marshal.PtrToStructure(mxlc.pamxctrl, GetType(MIXERCONTROL)), MIXERCONTROL)
            Else
                retValue = False
            End If
            Dim sizeofMIXERCONTROLDETAILS As Integer = Marshal.SizeOf(GetType(MIXERCONTROLDETAILS))
            Dim sizeofMIXERCONTROLDETAILS_UNSIGNED As Integer = Marshal.SizeOf(GetType(MIXERCONTROLDETAILS_UNSIGNED))
            pmxcd.cbStruct = sizeofMIXERCONTROLDETAILS
            pmxcd.dwControlID = mxc.dwControlID
            pmxcd.paDetails = Marshal.AllocCoTaskMem(sizeofMIXERCONTROLDETAILS_UNSIGNED)
            pmxcd.cChannels = 1
            pmxcd.item = 0
            pmxcd.cbDetails = sizeofMIXERCONTROLDETAILS_UNSIGNED

            rc = mixerGetControlDetailsA(hmixer, pmxcd, MIXER_GETCONTROLDETAILSF_VALUE)

            du = Marshal.PtrToStructure(pmxcd.paDetails, GetType(MIXERCONTROLDETAILS_UNSIGNED))

            vCurrentVol = du.dwValue

            Return retValue
        End If

        retValue = False
        Return retValue
    End Function 'GetVolumeControl

    Private Shared Function SetVolumeControl(ByVal hmixer As Integer, ByVal mxc As MIXERCONTROL, ByVal volume As Integer) As Boolean
        ' This function sets the value for a volume control. 
        ' Returns True if successful 
        Dim retValue As Boolean
        Dim rc As Integer
        Dim mxcd As New MIXERCONTROLDETAILS
        Dim vol As New MIXERCONTROLDETAILS_UNSIGNED

        mxcd.item = 0
        mxcd.dwControlID = mxc.dwControlID
        mxcd.cbStruct = Marshal.SizeOf(mxcd)
        mxcd.cbDetails = Marshal.SizeOf(vol)

        ' Allocate a buffer for the control value buffer 
        mxcd.cChannels = 1
        vol.dwValue = volume

        ' Copy the data into the control value buffer 
        mxcd.paDetails = Marshal.AllocCoTaskMem(Marshal.SizeOf(GetType(MIXERCONTROLDETAILS_UNSIGNED)))
        Marshal.StructureToPtr(vol, mxcd.paDetails, False)

        ' Set the control value 
        rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)

        If MMSYSERR_NOERROR = rc Then
            retValue = True
        Else
            retValue = False
        End If
        Return retValue
    End Function 'SetVolumeControl 

    Public Function GetVolume() As Integer
        Dim mixer As Integer
        Dim volCtrl As New MIXERCONTROL
        Dim currentVol As Integer
        mixerOpen(mixer, 0, 0, 0, 0) 'Returns the mixer control
        Dim type As Integer = MIXERCONTROL_CONTROLTYPE_VOLUME
        GetVolumeControl(mixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, type, volCtrl, currentVol)
        mixerClose(mixer)
        Return currentVol
    End Function 'GetVolume
    Public Function GetMax() As Integer
        Dim mixer As Integer
        Dim volCtrl As New MIXERCONTROL
        Dim currentVol As Integer
        mixerOpen(mixer, 0, 0, 0, 0)
        Dim type As Integer = MIXERCONTROL_CONTROLTYPE_VOLUME
        GetVolumeControl(mixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, type, volCtrl, currentVol)
        mixerClose(mixer)
        Return volCtrl.lMaximum
    End Function 'Gets max volume
    Public Function GetMuted() As String
        Dim mixer As Integer
        Dim volCtrl As New MIXERCONTROL
        Dim Muted As Integer
        mixerOpen(mixer, 0, 0, 0, 0)
        Dim type As Integer = MIXERCONTROL_CONTROLTYPE_MUTE
        GetVolumeControl(mixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, type, volCtrl, Muted)
        mixerClose(mixer)
        Return Muted
    End Function 'GetMuted status

    Public Sub SetVolume(ByVal vVolume As Integer)
        Dim mixer As Integer
        Dim volCtrl As New MIXERCONTROL
        Dim currentVol As Integer
        mixerOpen(mixer, 0, 0, 0, 0)
        Dim type As Integer = MIXERCONTROL_CONTROLTYPE_VOLUME
        GetVolumeControl(mixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, type, volCtrl, currentVol)
        If vVolume > volCtrl.lMaximum Then
            vVolume = volCtrl.lMaximum
        End If
        If vVolume < volCtrl.lMinimum Then
            vVolume = volCtrl.lMinimum
        End If
        SetVolumeControl(mixer, volCtrl, vVolume)
        GetVolumeControl(mixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, type, volCtrl, currentVol)
        If vVolume <> currentVol Then
            Throw New Exception("Cannot Set Volume")
        Else
            RaiseEvent VolumeChanged()
        End If
        mixerClose(mixer)
    End Sub 'SetVolume

    Public Sub SetMuted(ByVal boolMute As Boolean)
        ' This routine sets the volume setting of the current unit depending on the value passed through
        Dim mixer As Integer
        Dim volCtrl As New MIXERCONTROL
        Dim lngReturn As Integer
        Dim type As Integer = MIXERCONTROL_CONTROLTYPE_MUTE
        Dim currentVol As Integer
        ' Obtain the hmixer struct
        lngReturn = mixerOpen(mixer, 0, 0, 0, 0)
        ' Error check
        If lngReturn <> 0 Then Exit Sub
        ' Obtain the volumne control object
        GetVolumeControl(mixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, type, volCtrl, currentVol)
        ' Then set the volume
        SetVolumeControl(mixer, volCtrl, boolMute)
        mixerClose(mixer)
        RaiseEvent MuteChanged()
    End Sub 'Set the muted status

End Class
