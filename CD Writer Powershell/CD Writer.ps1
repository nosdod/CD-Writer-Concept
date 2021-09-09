#All-in-one:

param($pathToBurn = "Z:\HPE Work", $volumeName = "MyImage", $recorderIndex = 0, $closeMediaAfterBurn = $true)

$ok = $true

# Check the media in the drive
$dm = New-Object -ComObject "IMAPI2.MsftDiscMaster2"

#initialize the recorder:
$recorder = New-Object -ComObject "IMAPI2.MsftDiscRecorder2"

$recorder.InitializeDiscRecorder($dm.Item($recorderIndex))

Write-Output "INFO : Initialised recorder"
#use formatter to burn the data:

$df2d = New-Object -ComObject IMAPI2.MsftDiscFormat2Data
$df2d.Recorder = $recorder
$df2d.ClientName = "MyScriptBurner"
$df2d.ForceMediaToBeClosed = $closeMediaAfterBurn

if ( $df2d.MediaPhysicallyBlank -eq $false ) {
    Write-Output "ERROR : Disc must be blank"
    $ok = $false
}

if ( $ok ) {
    # Create the in memory disc image:
    $fsi = New-Object -ComObject "IMAPI2FS.MsftFileSystemImage"

    $fsi.FileSystemsToCreate = 7

    $fsi.VolumeName = $volumeName

    # Try to add the specified directory to the in memory file system
    Write-Output "INFO : Adding contents of $pathToBurn to burn image ..."
    try {
        $fsi.Root.AddTreeWithNamedStreams($pathToBurn, $false)
    } catch {
        Write-Output "Could not find path $pathToBurn : $PSItem"
        $ok = $false
    }

    if ( $ok ) {

        Write-Output "INFO : Creating iso disc image ... "
        $resultimage = $fsi.CreateResultImage()

        $resultStream = $resultimage.ImageStream

        Write-Output "INFO : Now burning the image ... "
        $df2d.Write($resultStream)
    }
}