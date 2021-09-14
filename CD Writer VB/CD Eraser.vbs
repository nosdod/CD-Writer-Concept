' CD Eraser
' 
' Erases CD-RW media
'
' Author : Mark Dodson (DodTech Ltd)
' Created : 09/09/2021

Option Explicit

'  Exit codes
Const E_NO_ERASE_POSSIBLE = -1
Const S_OK                =  0

' *** CD/DVD disc file system types
Const FsiFileSystemISO9660 = 1
Const FsiFileSystemJoliet  = 2
Const FsiFileSystemUDF102  = 4

' *** IFormat2Data Write Action Enumerations
Const IMAPI_FORMAT2_DATA_WRITE_ACTION_VALIDATING_MEDIA      = 0
Const IMAPI_FORMAT2_DATA_WRITE_ACTION_FORMATTING_MEDIA      = 1
Const IMAPI_FORMAT2_DATA_WRITE_ACTION_INITIALIZING_HARDWARE = 2
Const IMAPI_FORMAT2_DATA_WRITE_ACTION_CALIBRATING_POWER     = 3
Const IMAPI_FORMAT2_DATA_WRITE_ACTION_WRITING_DATA          = 4
Const IMAPI_FORMAT2_DATA_WRITE_ACTION_FINALIZATION          = 5
Const IMAPI_FORMAT2_DATA_WRITE_ACTION_COMPLETED             = 6
const IMAPI_FORMAT2_DATA_WRITE_ACTION_VERIFYING             = 7

' *** IMAPI2 Media Types 
Const IMAPI_MEDIA_TYPE_UNKNOWN            = 0  ' Media not present OR
                                               ' is unrecognized
Const IMAPI_MEDIA_TYPE_CDROM              = 1  ' CD-ROM
Const IMAPI_MEDIA_TYPE_CDR                = 2  ' CD-R
Const IMAPI_MEDIA_TYPE_CDRW               = 3  ' CD-RW
Const IMAPI_MEDIA_TYPE_DVDROM             = 4  ' DVD-ROM
Const IMAPI_MEDIA_TYPE_DVDRAM             = 5  ' DVD-RAM
Const IMAPI_MEDIA_TYPE_DVDPLUSR           = 6  ' DVD+R
Const IMAPI_MEDIA_TYPE_DVDPLUSRW          = 7  ' DVD+RW
Const IMAPI_MEDIA_TYPE_DVDPLUSR_DUALLAYER = 8  ' DVD+R dual layer
Const IMAPI_MEDIA_TYPE_DVDDASHR           = 9  ' DVD-R
Const IMAPI_MEDIA_TYPE_DVDDASHRW          = 10 ' DVD-RW
Const IMAPI_MEDIA_TYPE_DVDDASHR_DUALLAYER = 11 ' DVD-R dual layer
Const IMAPI_MEDIA_TYPE_DISK               = 12 ' Randomly writable

' *** IMAPI2 Data Media States
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_UNKNOWN            = 0
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_INFORMATIONAL_MASK = 15    
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_UNSUPPORTED_MASK   = 64512 '0xfc00
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_OVERWRITE_ONLY     = 1
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_BLANK              = 2
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_APPENDABLE         = 4     
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_FINAL_SESSION      = 8
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_DAMAGED            = 1024 '0x400
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_ERASE_REQUIRED     = 2048 '0x800
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_NON_EMPTY_SESSION  = 4096 '0x1000
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_WRITE_PROTECTED    = 8192 '0x2000
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_FINALIZED          = 16384 '0x4000
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_UNSUPPORTED_MEDIA  = 32768 '0x8000

Dim userPath:     userPath = ""
Dim isProduction: isProduction = True
Dim driveIndex:   driveIndex = 0
if WScript.Arguments.Count <> 0 Then
    if WScript.Arguments.Count >=1 Then
        userPath = WScript.Arguments.Item(0)
    end if

    if WScript.Arguments.Count >=2 Then
        isProduction = WScript.Arguments.Item(1)
    end if

    if WScript.Arguments.Count >=3 Then
        driveIndex = WScript.Arguments.Item(2)
    end if
end if

WScript.Quit(CDEraser())

Function CDEraser()
    Dim Index                ' Index to recording drive.
    Dim Recorder             ' Recorder object
    Dim Stream               ' Data stream for burning device
    Dim Ok                   ' Code is Ok to continue

    CDEraser = S_OK

    Ok = True
    
    Index = 0               ' Only disc writer on this system

    ' Create a DiscMaster2 object to connect to CD/DVD drives.
    WScript.Echo "INFO : Create connection to drive ..."
    Dim g_DiscMaster
    Set g_DiscMaster = WScript.CreateObject("IMAPI2.MsftDiscMaster2")

    ' Create a DiscRecorder object for the specified burning device.
    WScript.Echo "INFO : Create recorder interface ..."
    Dim uniqueId
    set recorder = WScript.CreateObject("IMAPI2.MsftDiscRecorder2")
    uniqueId = g_DiscMaster.Item(index)
    WScript.Echo "INFO : Initialise the recorder ..."
    recorder.InitializeDiscRecorder( uniqueId )

    Dim dataEraser    
            
    ' Define the new disc format and set the recorder
    WScript.Echo "INFO : Create an eraser object ..."
    Set dataEraser = CreateObject ("IMAPI2.MsftDiscFormat2Erase")
    dataEraser.recorder = Recorder
    dataEraser.ClientName = "IMAPIv2 Eraser"

    Dim mediaType
    if isMediaTypeAvailable(mediaType,dataEraser) Then
        WScript.Echo "Current Media Type"
        DisplayMediaType(mediaType)
    else
        Ok = False
    end if

    if Ok Then
        ' Attach event handler to the data writing object.
        WScript.ConnectObject  dataEraser, "dwBurnEvent_"

        WScript.Echo "Erasing the disc..."
        dataEraser.EraseMedia()

        WScript.Echo "----- Finished erasing media -----"
    end if

    if Ok <> True Then
        CDWriter = E_NO_BURN_POSSIBLE
    end if
End Function

' Check the type of media loaded - if possible
Function isMediaTypeAvailable(ByRef mediaType, dataEraser)
    On Error Resume Next

    isMediaTypeAvailable = True    

    ' Must catch an exception here which happens if no media is present
    mediaType = dataEraser.CurrentPhysicalMediaType
    if Err.Number <> 0 Then
        WScript.Echo Err.Description
        isMediaTypeAvailable = False
    end if
End Function

' Event handler - Progress updates when writing data
Sub dwBurnEvent_Update( byRef object, byRef elapsedSeconds, byRef estimatedTotalSeconds )
    Dim strTimeStatus
    strTimeStatus = "Time: " & elapsedSeconds & _
        " / " & estimatedTotalSeconds
   
    WScript.Echo "Progress:  " & strTimeStatus
End Sub

Sub DisplayMediaType(dMediaType)
    Select Case dmediaType 
        Case IMAPI_MEDIA_TYPE_UNKNOWN
            WScript.Echo "    Empty device or an unknown disc type."
        
        Case IMAPI_MEDIA_TYPE_CDROM
            WScript.Echo "    CD-ROM"
        
        Case IMAPI_MEDIA_TYPE_CDR
            WScript.Echo "    CD-R"
        
        Case IMAPI_MEDIA_TYPE_CDRW
            WScript.Echo "    CD-RW"
        
        Case IMAPI_MEDIA_TYPE_DVDROM
            WScript.Echo "    Read-only DVD drive and/or disc"
        
        Case IMAPI_MEDIA_TYPE_DVDRAM
            WScript.Echo "    DVD-RAM"
        
        Case IMAPI_MEDIA_TYPE_DVDPLUSR
            WScript.Echo "    DVD+R"
        
        Case IMAPI_MEDIA_TYPE_DVDPLUSRW
            WScript.Echo "    DVD+RW"
        
        Case IMAPI_MEDIA_TYPE_DVDPLUSR_DUALLAYER
            WScript.Echo "    DVD+R Dual Layer media"
        
        Case IMAPI_MEDIA_TYPE_DVDDASHR
            WScript.Echo "    DVD-R"
        
        Case IMAPI_MEDIA_TYPE_DVDDASHRW
            WScript.Echo "    DVD-RW"
        
        Case IMAPI_MEDIA_TYPE_DVDDASHR_DUALLAYER
            WScript.Echo "    DVD-R Dual Layer media"
        
        Case IMAPI_MEDIA_TYPE_DISK
            WScript.Echo "    Randomly-writable, hardware-defect " _
                + "managed media type "
            WScript.Echo "    that reports the ""Disc"" profile " _
                + "as current."
        End Select
End Sub