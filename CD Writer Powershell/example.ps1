#Create file system object

$fsi = New-Object -ComObject IMAPI2FS.MsftFileSystemImage
$fsi.FileSystemsToCreate = 7
$fsi.VolumeName = "MyImage"

# Add files to your file system:
# Z: is my Synology212 NAS device, HPE Work is a directory on it
$fsi.Root.AddTreeWithNamedStreams("Z:\HPE Work")
$resultimage = $fsi.CreateResultImage()
$resultStream = $resultimage.ImageStream

# Enumerate available recorders through MsftDiscMaster2
$dm = New-Object -ComObject IMAPI2.MsftDiscMaster2
$dm
# On my machine I only have one disc writer, if you have more than one you need to note the position
# in the list (starting at 0) of the device you want to record to for use later.
# Create DiscRecorder object
$recorder = New-Object -ComObject IMAPI2.MsftDiscRecorder2
# Initialize recorder with unique id from discmaster2

# In the next command use the position in the list in $dm.Item(..)
$recorder.InitializeDiscRecorder($dm.Item(0))

# And Formatter, which will be used to format you filesystem image in appropriate way for recording onto your media and write data using your recorder:
$df2d = New-Object -ComObject IMAPI2.MsftDiscFormat2Data
$df2d.Recorder = $recorder

# Put client name, which will be shown for other applications, once you start writing
$df2d.ClientName = "MyScriptBurner" 
# Now you set up everything needed, and can start burning, just call Write from DiscFormat2Data:
$df2d.Write($resultStream)