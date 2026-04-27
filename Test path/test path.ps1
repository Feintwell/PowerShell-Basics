$Folder = 'C:\Users\User\Code'
"Test to see if folder [$Folder]  exists"
if (Test-Path -Path C:\Windows) {
    "Path exists!"
} else {
    "Path doesn't exist."
}