#All-in-one:

param($pathToBurn = "Z:\HPE Work 2", $recorderIndex = 0, $closeMediaAfterBurn = $true)

$ok = $true

#create the image:

$fsi = New-Object -ComObject "IMAPI2FS.MsftFileSystemImage"

$fsi.FileSystemsToCreate = 7

$fsi.VolumeName = "MyImage"

# Try to add the specified directory to the in memory file system
try {
    $fsi.Root.AddTreeWithNamedStreams($pathToBurn, $false)
} catch {
    Write-Output "Could not find path $pathToBurn : $PSItem"
    $ok = $false
}

if ( $ok ) {

    $resultimage = $fsi.CreateResultImage()

    $resultStream = $resultimage.ImageStream

    #initialize the recorder:

    $dm = New-Object -ComObject "IMAPI2.MsftDiscMaster2"

    $recorder = New-Object -ComObject "IMAPI2.MsftDiscRecorder2"

    $recorder.InitializeDiscRecorder($dm.Item($recorderIndex))

    #use formatter to burn the data:

    $df2d = New-Object -ComObject IMAPI2.MsftDiscFormat2Data

    $df2d.Recorder = $recorder

    $df2d.ClientName = "MyScriptBurner"

    $df2d.ForceMediaToBeClosed = $closeMediaAfterBurn

    #burn it:

    $df2d.Write($resultStream)
}