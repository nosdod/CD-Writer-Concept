The code here was taken from https://docs.microsoft.com/en-us/archive/blogs/opticalstorage/writing-optical-discs-using-imapi-2-in-powershell

Initial experiments were done using a Pioneer Blue Ray writer connected to an Alienware (Dell) laptop running Windws 10 pro
No additional drivers etc needed installing for the powershell commands to run.

The script worked when a blank CD-RW was present in the drive.

** Observations
When run a second time with the same CD-RW in the drive, the script completes quickly but leaves the drive doing something!
I suspect Windows has detected the disc isn't blank and enforces a format operation ?
CD-RW Supports multi-session writes - so each new iso is probably written as a new session - will check for a blank disc for now.
This isn't really relevent for this application as CD-R media will be used - hence they will ALWAYS be blank.
It needs to be handled though as we want to use CD-RW media during developemt so as to not waste CD-R media.

Adding the source files to the in memory file system took a while - will need to provide user feedback during this operation.

Writing to the disc takes some time - will need to provide user feedback during this operation.

There is no error handling in the example script.

The example script doesn't close the disc.