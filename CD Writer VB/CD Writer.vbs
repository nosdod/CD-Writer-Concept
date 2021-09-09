' CD Writer
' 
' Writes all files in a given directory to blank media
' In production environments only CD-R media is accepted
' Any writeable media is permitted in other environments.
' 
' Uses IMAPI v2 to talk to the specified device
'
' If write succeeds an exit code of 0 is returned
' 
' Author : Mark Dodson (DodTech Ltd)
' Created : 09/09/2021

Option Explicit

'  Exit codes
Const E_NO_BURN_POSSIBLE = -1
Const S_OK               =  0

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
Const IMAPI_FORMAT2_DATA_MEDIA_STATE_UNSUPPORTED_MASK   = 61532 '0xfc00
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
if WScript.Arguments.Count <> 0 Then
    if WScript.Arguments.Count >=1 Then
        userPath = WScript.Arguments.Item(0)
    end if

    if WScript.Arguments.Count >=2 Then
        isProduction = WScript.Arguments.Item(1)
    end if
end if

WScript.Quit(CDWriter(userPath,isProduction))

Function CDWriter(userPath,isProduction)
    Dim Index                ' Index to recording drive.
    Dim Recorder             ' Recorder object
    Dim Path                 ' Directory of files to burn
    Dim Stream               ' Data stream for burning device
    Dim Ok                   ' Code is Ok to continue

    CDWriter = S_OK

    Ok = True
    
    Index = 0               ' Only disc writer on this system

    if Len(userPath) = 0 Then
        Path = "Z:\HPE Work"    ' Files to transfer to disc
    else
        Path = userPath
    end if 

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

    ' Create an image stream for a specified directory.
    Dim FSI                  'Disc file system
    Dim Dir                  'Root directory of the disc file system
    Dim dataWriter    
            
    ' Create a new file system image and retrieve root directory
    WScript.Echo "INFO : Create an in memory file system ..."
    Set FSI = CreateObject("IMAPI2FS.MsftFileSystemImage")
    Set Dir = FSI.Root

    ' Define the new disc format and set the recorder
    WScript.Echo "INFO : Create a formatter object ..."
    Set dataWriter = CreateObject ("IMAPI2.MsftDiscFormat2Data")
    dataWriter.recorder = Recorder
    dataWriter.ClientName = "IMAPIv2 TEST"

    Dim isRecorderSupported
    isRecorderSupported = dataWriter.IsRecorderSupported(recorder)
    If isRecorderSupported then
        WScript.Echo "--- Current recorder IS supported. ---"
    else
        WScript.Echo "Current recorder IS NOT supported."
        Ok = False
    end if

    Dim isMediaSupported
    isMediaSupported = dataWriter.IsCurrentMediaSupported(recorder)

    If isMediaSupported then
        WScript.Echo "--- Current media IS supported. ---"
    else
        WScript.Echo "Current media IS NOT supported."
        Ok = False
    end if

    ' Must catch an exception here which happens if no media is present
    Dim mediaType
    if isMediaTypeAvailable(mediaType,dataWriter) Then
        WScript.Echo "Current Media Type"
        DisplayMediaType(mediaType)

        ' Check a few CurrentMediaStatus possibilities. Each status
        ' is associated with a bit and some combinations are legal.
        WScript.Echo "Checking Current Media Status"        
        Dim curMediaStatus
        curMediaStatus = dataWriter.CurrentMediaStatus
        DisplayMediaStatus(curMediaStatus)
    else
        Ok = False
    end if

    if Ok Then
        if mediaType <> IMAPI_MEDIA_TYPE_CDR Then
            if isProduction Then
                WScript.Echo "ERROR : In production environments only CD-R media is supported"
                Ok = False
            else
                WScript.Echo "WARNING : In production environments only CD-R media is supported"
            end if
        end if

        if Ok Then
            WScript.Echo "INFO : Check disc is blank ..."
            if dataWriter.MediaPhysicallyBlank Then
                FSI.FreeMediaBlocks = dataWriter.FreeSectorsOnMedia
                FSI.FileSystemsToCreate = FsiFileSystemISO9660

                ' Add the directory and its contents to the file system 
                WScript.Echo "INFO : Add the source files to the in memory file system ..."
                Dir.AddTree Path, false
                
                ' Create an image from the file system
                WScript.Echo "INFO : Build the ISO image of the filesystem ..."
                Dim result
                Set result = FSI.CreateResultImage()
                Stream = result.ImageStream
            
                ' Attach event handler to the data writing object.
                WScript.ConnectObject  dataWriter, "dwBurnEvent_"

                ' Specify the recorder and write the stream to disc.
                WScript.Echo "Writing the ISO image to disc..."
                dataWriter.write(Stream)

                WScript.Echo "----- Finished writing content -----"
                Main = 0
            else 
                WScript.Echo "ERROR : Disc must be blank"
            end if
        end if
    end if

    if Ok <> True Then
        CDWriter = E_NO_BURN_POSSIBLE
    end if
End Function

' Check the type of media loaded - if possible
Function isMediaTypeAvailable(ByRef mediaType, dataWriter)
    On Error Resume Next

    isMediaTypeAvailable = True    

    ' Must catch an exception here which happens if no media is present
    mediaType = dataWriter.CurrentPhysicalMediaType
    if Err.Number <> 0 Then
        WScript.Echo Err.Description
        isMediaTypeAvailable = False
    end if
End Function

' Event handler - Progress updates when writing data
SUB dwBurnEvent_Update( byRef object, byRef progress )
    DIM strTimeStatus
    strTimeStatus = "Time: " & progress.ElapsedTime & _
        " / " & progress.TotalTime
   
    SELECT CASE progress.CurrentAction
    CASE IMAPI_FORMAT2_DATA_WRITE_ACTION_VALIDATING_MEDIA
        WScript.Echo "Validating media " & strTimeStatus

    CASE IMAPI_FORMAT2_DATA_WRITE_ACTION_FORMATTING_MEDIA
        WScript.Echo "Formatting media " & strTimeStatus
        
    CASE IMAPI_FORMAT2_DATA_WRITE_ACTION_INITIALIZING_HARDWARE
        WScript.Echo "Initializing Hardware " & strTimeStatus

    CASE IMAPI_FORMAT2_DATA_WRITE_ACTION_CALIBRATING_POWER
        WScript.Echo "Calibrating Power (OPC) " & strTimeStatus

    CASE IMAPI_FORMAT2_DATA_WRITE_ACTION_WRITING_DATA
        DIM totalSectors, writtenSectors, percentDone
        totalSectors = progress.SectorCount
        writtenSectors = progress.LastWrittenLba - progress.StartLba
        percentDone = FormatPercent(writtenSectors/totalSectors)
        WScript.Echo "Progress:  " & percentDone & "  " & strTimeStatus

    CASE IMAPI_FORMAT2_DATA_WRITE_ACTION_FINALIZATION
        WScript.Echo "Finishing the writing " & strTimeStatus
    
    CASE IMAPI_FORMAT2_DATA_WRITE_ACTION_COMPLETED
        WScript.Echo "Completed the burn."

    CASE IMAPI_FORMAT2_DATA_WRITE_ACTION_VERIFYING
        WScript.Echo "Verifying the data."

    CASE ELSE
        WScript.Echo "Unknown action: " & progress.CurrentAction
    END SELECT
END SUB

Sub DisplayMediaStatus(mediaStatus)
    if IMAPI_FORMAT2_DATA_MEDIA_STATE_FINALIZED AND mediaStatus then
        WScript.Echo "    Media has already been finalised."
    end if

    if IMAPI_FORMAT2_DATA_MEDIA_STATE_UNKNOWN AND mediaStatus then
        WScript.Echo "    Media state is unknown."
    end if

    if IMAPI_FORMAT2_DATA_MEDIA_STATE_OVERWRITE_ONLY AND mediaStatus then
        WScript.Echo "    Currently, only overwriting is supported."
    end if

    if IMAPI_FORMAT2_DATA_MEDIA_STATE_APPENDABLE AND mediaStatus then
        WScript.Echo "    Media is currently appendable."
    end if

    if IMAPI_FORMAT2_DATA_MEDIA_STATE_FINAL_SESSION AND mediaStatus then
        WScript.Echo "    Media is in final writing session."
    end if

    if IMAPI_FORMAT2_DATA_MEDIA_STATE_DAMAGED AND mediaStatus then
        WScript.Echo "    Media is damaged."
    end if

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