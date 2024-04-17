$path = [Environment]::GetFolderPath("MyPictures") + "\iPhone Fotos"

Add-Type -AssemblyName System.Runtime.WindowsRuntime

[void][Windows.Foundation.IAsyncOperation`1,Windows.Foundation,ContentType=WindowsRuntime]
[void][Windows.Foundation.IAsyncOperationWithProgress`2,Windows.Foundation,ContentType=WindowsRuntime]
[void][Windows.Media.Import.PhotoImportManager,Windows.Media.Import,ContentType=WindowsRuntime]


# Await Function 1
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | 
    Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and 
                    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

Function Await($WinRtTask, $ResultType)
{
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}

# Present the available devices in a basic GUI, for user to choose one
$sources = Await (
                [Windows.Media.Import.PhotoImportManager]::FindAllSourcesAsync()
                ) (
                [System.Collections.Generic.IReadOnlyList[Windows.Media.Import.PhotoImportSource]]
                )

$selectedSource = $sources | Select-Object DisplayName, Model, Manufacturer, ConnectionTransport, ConnectionProtocol, PowerSource, SerialNumber | Out-GridView -OutputMode Single -Title 'Kies een apparaat'

foreach ($source in $sources) {
    if ($selectedSource.SerialNumber -eq $source.SerialNumber) {
        $selectedSource = $source
    }
}

if ($selectedSource -eq $null) {
    Write-OUtput "-----------------------------------------------------------------"
    Write-OUtput "| Verbind uw iPhone eerst via een USB-kabel met uw computer     |"
    Write-OUtput "| Ten tweede, geef je via uw iPhone toegang aan deze computer   |"
    Write-OUtput "| Hierna open je dit programma opnieuw                          |"
    Write-OUtput "| Ten slotte kies je opnieuw een apparaat in het programma      |"
    Write-OUtput "-----------------------------------------------------------------"
    Start-Sleep -Seconds 3
    exit
}

If (-not (Test-Path $path)) {
    New-Item -Path $path -ItemType Directory | Out-Null
}

$folder = Await (
                [Windows.Storage.StorageFolder]::GetFolderFromPathAsync($path)
                ) (
                [Windows.Storage.StorageFolder]
                )

# Start an import session for the device
$importSession = $selectedSource.CreateImportSession()
$importSession.DestinationFolder = $folder
$importSession.AppendSessionDateToDestinationFolder = $false


# Await Function 2
$asTaskGeneric2 = ([System.WindowsRuntimeSystemExtensions].GetMethods() | 
    Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and 
                    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperationWithProgress`2' })[0]

Function Await2($WinRtTask, $ResultType1, $ResultType2)
{
    $asTask = $asTaskGeneric2.MakeGenericMethod($ResultType1, $ResultType2)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}

# Find all files on the device
$items = Await2 ( $importSession.FindItemsAsync(
                [Windows.Media.Import.PhotoImportContentTypeFilter]::ImagesAndVideos, 
                [Windows.Media.Import.PhotoImportItemSelectionMode]::SelectNone)
                ) (
                [Windows.Media.Import.PhotoImportFindItemsResult]
                ) (
                [uint32]
                )

if ($items.PhotosCount -eq 0) {
    Write-Output "---------------------------------------------------------------------------------------------------"
    Write-Output "Het programma is gefaald aangezien u nog geen toegang heeft gegeven tot uw bestanden"
    Write-OUtput "Open eerst uw iPhone om toegang te geven!"
    Write-Output "---------------------------------------------------------------------------------------------------"
    Start-Sleep -Seconds 3
    exit  
}

$files = Get-ChildItem -Path $path -File
$current_items = $files | Sort-Object LastWriteTime[0] -Descending
$updated_items = $items.FoundItems | Sort-Object Date.DateTime[0] -Descending

# Checks if new files have been added, old aren't re-transferred
foreach ($item in $updated_items) {
    if ($item.Name -notin $current_items.Name) {
        $item.IsSelected = $true
    }
}

Write-Output "---------------------------------------------------------------------------------------------------"
Write-Output "Uw bestanden worden ge�mporteerd, dit programma niet afsluiten..."    

# Run the import and wait for it to finish
$importResult = Await2 ( $items.ImportItemsAsync()
                        ) (
                        [Windows.Media.Import.PhotoImportImportItemsResult]
                        ) (
                        [Windows.Media.Import.PhotoImportProgress]
                        )

if ($importResult.TotalCount -ne 0) {
    Write-Output "---------------------------------------------------------------------------------------------------"
    Write-Output "Er zijn $($importResult.PhotosCount) nieuwe foto's ge�mporteerd"
    Write-Output "Er zijn $($importResult.VideosCount) nieuwe video's ge�mporteerd"
    Write-Output "---------------------------------------------------------------------------------------------------"
    Write-Output "Het totaal aantal toegevoegde bestanden is: $($importResult.TotalCount)"
    Start-Sleep -Seconds 2
    Start Explorer $path
    Start-Sleep -Seconds 2
} else {
    Write-Output "---------------------------------------------------------------------------------------------------"
    Write-Output "Er hoeven geen nieuwe bestanden meer toegevoegd te worden, uw fotoalbum is al reeds compleet"
    Write-Output "---------------------------------------------------------------------------------------------------"
    Start-Sleep -Seconds 3
}