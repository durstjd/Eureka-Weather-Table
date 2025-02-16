Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition '
public class DPIAware
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
'

[System.Windows.Forms.Application]::EnableVisualStyles()
[void] [DPIAware]::SetProcessDPIAware()

class WeatherJson {
    [string]$jTime
    [string]$jWeather
    [string]$jEorzeanTime
}

class EorzeaWeatherTable {
    [string]$Zone
    [string[]]$Weather
    [int32[]]$Chance
}

# Initialize an array to hold all weather table entries
$weatherTable = @()

# Create and populate the weather table for each zone
$weatherTable += [EorzeaWeatherTable]@{
    Zone   = 'Anemos'
    Chance = @(30, 30, 30, 10)
    Weather = @('Fair Skies', 'Gales', 'Showers', 'Snow')
}

$weatherTable += [EorzeaWeatherTable]@{
    Zone   = 'Pagos'
    Chance = @(10, 18, 18, 18, 18, 18)
    Weather = @('Clear Skies', 'Fog', 'Heat Waves', 'Snow', 'Thunder', 'Blizzard')
}

$weatherTable += [EorzeaWeatherTable]@{
    Zone   = 'Pyros'
    Chance = @(10, 18, 18, 18, 18, 18)
    Weather = @('Fair Skies', 'Heat Waves', 'Thunder', 'Blizzard', 'Umbral Wind', 'Snow')
}

$weatherTable += [EorzeaWeatherTable]@{
    Zone   = 'Hydatos'
    Chance = @(12, 22, 22, 22, 22)
    Weather = @('Fair Skies', 'Showers', 'Gloom', 'Thunderstorms', 'Snow')
}

#Eorzea Constant -- 24 Earth hours * 60 Earth minutes / 70 (Eorzea minutes in a day)
$eorzeaConstant = 24*60/70

#Earth Epoch Time
$epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (Get-Date).ToUniversalTime()).TotalSeconds)

#Eorzea Time -- The product of the epoch time and Eorzea Constant
$eorzeaTime = ([datetime]"01/01/1970 00:00:00").AddSeconds($epoch * $eorzeaConstant)

<#
function CalculateEorzeaTime {
    param (
        [Parameter(Mandatory = $true)]
        [double]$unixSeconds
    )
    return [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (([datetime]"01/01/1970 00:00:00").AddSeconds($unixSeconds * $eorzeaConstant)).ToUniversalTime()).TotalSeconds)
}
#>

function CalculateEorzeaTime {
    param (
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$inputTime
    )

    $unixSeconds = 0
    if (!$inputTime){$inputTime = $(Get-Date)}


    if ($inputTime -is [double]) {
        $unixSeconds = $inputTime
    } elseif ($inputTime -is [datetime]) {
        $unixSeconds = [Math]::Floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End ($inputTime.ToUniversalTime())).TotalSeconds)
    } else {
        throw "The input must be a double or a DateTime."
    }

    if ($unixSeconds -lt 0) {
        throw "Unix timestamp cannot be negative."
    }

    $adjustedEorzeaTime = $unixSeconds * $eorzeaConstant

    return [Math]::Floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (([datetime]"01/01/1970 00:00:00").AddSeconds($adjustedEorzeaTime)).ToUniversalTime()).TotalSeconds)
}

#Calculate Weather Target. Use the real world time ($epoch). 
#Subtract the $weatherTable[$index].Chance[$chanceIndex] (incremented) from $weatherTarget until it is less than 0. Then, $weatherTable[$index].Weather[$chanceIndex] is the current weather.
#This can be forecasted by adding or subtracting 23m20s.
#This should always fall within a range EorzeaTime 0:00-07:59, 08:00-15:59, 16:00-23:59. 
function CalculateTarget {
    param (
        [Parameter(Mandatory = $true)]
        [double]$unixSeconds
    )

    # Get Eorzea hour for weather start
    $bell = [math]::Floor($unixSeconds / 175)

    # For these calculations, 16:00 is 0, 00:00 is 8 and 08:00 is 16.
    # This rotates the time accordingly and holds onto the leftover time
    # to add back in later on.
    $increment = ([uint32](($bell + 8) - ($bell % 8))) % 24

    # Take Eorzea days since the Unix epoch
    $totalDays = [uint32]([math]::Floor($unixSeconds / 4200))

    # Calculate the base value
    $calcBase = ($totalDays * 0x64) + $increment

    # Calculate the weather interval with the next two bitwise operations
    # Perform the first bitwise operation
    $step1 = ($calcBase -shl 0xB) -bxor $calcBase

    # Perform the second bitwise operation
    $step2 = ($step1 -shr 8) -bxor $step1

    # Return the final result
    return [int32]($step2 % 0x64)
}

# Function to get the current weather based on the given zone and real-world time
function GetCurrentWeather {
    param (
        [string]$zone,
        [int]$weatherTarget
    )

    # Find the index corresponding to the given zone
    $index = $weatherTable.Zone.IndexOf($zone)
    if ($index -eq -1) {
        throw "Zone '$zone' not found in the weather table."
    }

    # Initialize variables
    $currentChanceIndex = 0
    $remainingWeatherTarget = $weatherTarget

    # Loop to find the current weather
    while ($remainingWeatherTarget -ge 0) {
        $remainingWeatherTarget -= $weatherTable[$index].Chance[$currentChanceIndex]
        $currentChanceIndex++
    }

    # Return the weather at the calculated chance index
    return $weatherTable[$index].Weather[$currentChanceIndex - 1]
}

function CalculateIncrement {
    param (
        [Parameter(Mandatory = $true)]
        [double]$unixSeconds
    )

    # Get Eorzea hour for weather start
    $bell = [math]::Floor($unixSeconds / 175)

    # For these calculations, 16:00 is 0, 00:00 is 8 and 08:00 is 16.
    # This rotates the time accordingly and holds onto the leftover time
    # to add back in later on.
    $increment = ([uint32](($bell + 8) - ($bell % 8))) % 24

    if ($increment -eq 0){
        return "16:00 - 23:59"
    }
    elseif ($increment -eq 8) {
        return "00:00 - 07:59"
    }
    else {
        return "08:00 - 15:59"
    }
}


<# NOTES
Forecasting.

Change the end datetime to the target. Do all in local time or all in Zulu time.

$epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End ([datetime]"02/14/2025 13:24").ToUniversalTime()).TotalSeconds)
GetCurrentWeather -zone "Pyros" -weatherTarget $(CalculateTarget -unixSeconds $epoch)



This gives us our real life minutes into an Eorzean day. 70 minutes = 1 day (bell).
23.333333333 minutes between the 8hr periods of the day.

($([Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (Get-Date).ToUniversalTime()).TotalSeconds))/60)%70

Now lets get when this window started in real-life time, and we can adjust that by increments of 23.3333333 minutes (23 minutes, 20 seconds) to get any time period relative to now.
We'll go back 2 hours (6 windows) to monitor when the last time a Eureka NM could spawn.

$val = ($([Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (Get-Date).ToUniversalTime()).TotalSeconds))/60)%70
(Get-Date).AddMinutes(($val%23.333333333)*-1)



#>

function GetForecast {
    param (
        [string]$zone
    )

    Write-Host "Weather in"$zone
    Write-Host "Current Eorzea Time: "$eorzeaTime.ToString("HH:mm")
    #Previous Weather
    $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End ($(Get-Date).AddMinutes(-23.33333333)).ToUniversalTime()).TotalSeconds)
    $weather = GetCurrentWeather -zone $zone -weatherTarget $(CalculateTarget -unixSeconds $epoch)
    $thisIncrement = CalculateIncrement -unixSeconds $epoch
    Write-Host "Previous Weather: "$weather" from "$thisIncrement

    #Current Weather
    $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (Get-Date).ToUniversalTime()).TotalSeconds)
    $weather = GetCurrentWeather -zone $zone -weatherTarget $(CalculateTarget -unixSeconds $epoch)
    $thisIncrement = CalculateIncrement -unixSeconds $epoch
    Write-Host "Current Weather: "$weather" from "$thisIncrement

    #Next Weather
    $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End ($(Get-Date).AddMinutes(23.33333333)).ToUniversalTime()).TotalSeconds)
    $weather = GetCurrentWeather -zone $zone -weatherTarget $(CalculateTarget -unixSeconds $epoch)
    $thisIncrement = CalculateIncrement -unixSeconds $epoch
    Write-Host "Upcoming Weather: "$weather" from "$thisIncrement

}

function GetAdjustedForecast {
    param (
        [string]$zone
    )
    # Our weather interval's start time.
    $adjustedTime = ($([Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (Get-Date).ToUniversalTime()).TotalSeconds))/60)%70
    $intervalStart = (Get-Date).AddMinutes(($adjustedTime%23.333333333)*-1)
    Write-Host $intervalStart

    Write-Host "Weather in"$zone
    Write-Host "Current Eorzea Time: "$eorzeaTime.ToString("HH:mm")
    #Previous Weather
    $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End ($intervalStart.AddMinutes(-23.33333333)).ToUniversalTime()).TotalSeconds)
    $weather = GetCurrentWeather -zone $zone -weatherTarget $(CalculateTarget -unixSeconds $epoch)
    $thisIncrement = CalculateIncrement -unixSeconds $epoch
    Write-Host "Previous Weather: "$weather" from "$thisIncrement "("$intervalStart.AddMinutes(-23.33333333)")"

    #Current Weather
    $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End $intervalStart.ToUniversalTime()).TotalSeconds)
    $weather = GetCurrentWeather -zone $zone -weatherTarget $(CalculateTarget -unixSeconds $epoch)
    $thisIncrement = CalculateIncrement -unixSeconds $epoch
    Write-Host "Current Weather: "$weather" from "$thisIncrement "("$intervalStart")"

    #Next Weather
    $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End ($intervalStart.AddMinutes(23.33333333)).ToUniversalTime()).TotalSeconds)
    $weather = GetCurrentWeather -zone $zone -weatherTarget $(CalculateTarget -unixSeconds $epoch)
    $thisIncrement = CalculateIncrement -unixSeconds $epoch
    Write-Host "Upcoming Weather: "$weather" from "$thisIncrement "("$intervalStart.AddMinutes(23.33333333)")"

}

function GetWeatherData {
    param (
        [Parameter(Mandatory = $true)]
        [string]$zone
    )
    # Our weather interval's start time.
    $adjustedTime = ($([Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (Get-Date).ToUniversalTime()).TotalSeconds))/60)%70
    $intervalStart = (Get-Date).AddMinutes(($adjustedTime%23.333333333)*-1)
    Write-Host "Current Time Window: "$intervalStart

    # Initialize $thisIncrement
    $thisInterval = -6

    # Use a for loop to iterate through $thisIncrement from -6 to 18
    for ($thisInterval = -6; $thisInterval -le 18; $thisInterval++) {
        $interval = $intervalStart.AddMinutes(23.333333333*(-1+$thisInterval))
        $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End $interval.ToUniversalTime()).TotalSeconds)
        $weather = GetCurrentWeather -zone $zone -weatherTarget $(CalculateTarget -unixSeconds $epoch)
        $thisIncrement = CalculateIncrement -unixSeconds $epoch
        Write-Host $interval" : "$weather" from "$thisIncrement
    }
}

#GetWeatherData -zone Pyros

function GetIntervalStart {
    $adjustedTime = ($([Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End (Get-Date).ToUniversalTime()).TotalSeconds))/60)%70
    return (Get-Date).AddMinutes(($adjustedTime%23.333333333)*-1)
}

function GetWeatherDataJson {
    param (
        [Parameter(Mandatory = $true)]
        [string]$zone
    )

    # Our array to hold the weather data we generate
    $weatherData = @()

    # Our weather interval's start time.
    $intervalStart = GetIntervalStart

    # Initialize $thisIncrement
    $thisInterval = -5
    $thisEndInterval = 21
    $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End $intervalStart.ToUniversalTime()).TotalSeconds)
    $thisIncrement = CalculateIncrement -unixSeconds $epoch
    if ($thisIncrement -eq "08:00 - 15:59"){$thisInterval -= 1; $thisEndInterval -= 1}
    if ($thisIncrement -eq "16:00 - 23:59"){$thisInterval -= 2; $thisEndInterval -= 2}

    # Use a for loop to iterate through $thisIncrement from -6 to 18
    for ($thisInterval; $thisInterval -le $thisEndInterval; $thisInterval++) {
        $interval = $intervalStart.AddMinutes(23.333333333*(-1+$thisInterval))
        $epoch = [Math]::floor((New-TimeSpan -Start ([datetime]"01/01/1970 00:00:00") -End $interval.ToUniversalTime()).TotalSeconds)
        $weather = GetCurrentWeather -zone $zone -weatherTarget $(CalculateTarget -unixSeconds $epoch)
        $thisIncrement = CalculateIncrement -unixSeconds $epoch
        $weatherData += [WeatherJson]@{
            jTime = $interval.ToString("HH:mm")
            jWeather = $weather
            jEorzeanTime = $thisIncrement
        }
    }
    return ($weatherData | ConvertTo-Json)
}

function WeatherAlarm {
    $alarm = New-Object System.Media.SoundPlayer
    $alarm.SoundLocation = "$PSScriptRoot\assets\yoooooo.wav"
    $alarm.PlaySync()
}

function DisplayWeatherTable {
    param (
        [Parameter(Mandatory = $true)]
        [string]$zone
    )

    # JSON data
    $jsonData = GetWeatherDataJson -zone $zone

    # Adjusted time (example value)
    $highlightedTime = (GetIntervalStart).ToString("HH:mm")

    # Convert JSON to PowerShell objects
    $data = $jsonData | ConvertFrom-Json

    # Define the fixed order of Eorzean time periods
    $eorzeanTimes = @("00:00 - 07:59", "08:00 - 15:59", "16:00 - 23:59")

    # Group data into rows
    <#$rows = @()
    $currentRow = @{}
    foreach ($entry in $data) {
        $currentRow[$entry.jEorzeanTime] = "$($entry.jTime) - $($entry.jWeather)"
    
        # If all Eorzean times are filled, finalize the row
        if ($currentRow.Count -eq $eorzeanTimes.Count) {
            $rows += $currentRow
            $currentRow = @{}
        }
    }

    # Add any remaining partial row (if data doesn't fill all columns)
    if ($currentRow.Count -gt 0) {
        $rows += $currentRow
    }#>
    function Group-DataIntoRows {
        param (
            [Parameter(Mandatory = $true)]
            $data
        )
        $rows = @()
        $currentRow = @{}
        foreach ($entry in $data) {
            $currentRow[$entry.jEorzeanTime] = "$($entry.jTime) - $($entry.jWeather)"
            
            # If all Eorzean times are filled, finalize the row
            if ($currentRow.Count -eq $eorzeanTimes.Count) {
                $rows += $currentRow
                $currentRow = @{}
            }
        }

        # Add any remaining partial row (if data doesn't fill all columns)
        if ($currentRow.Count -gt 0) {
            $rows += $currentRow
        }

        return $rows
    }
    # Create a DataTable to hold the data for the DataGridView
    $dataTable = New-Object System.Data.DataTable

    # Add columns to the DataTable in the fixed order
    foreach ($et in $eorzeanTimes) {
        $dataTable.Columns.Add($et) | Out-Null
    }

    # Populate the DataTable with rows
    <#foreach ($row in $rows) {
        $rowData = @()
        foreach ($et in $eorzeanTimes) {
            if ($row.ContainsKey($et)) {
                $rowData += $row[$et]
            } else {
                $rowData += ""
            }
        }
        $dataTable.Rows.Add($rowData) | Out-Null
    }#>
    function Populate-DataTable {
        param (
            [Parameter(Mandatory = $true)]
            $rows
        )
        $dataTable.Rows.Clear() # Clear existing rows
        foreach ($row in $rows) {
            $rowData = @()
            foreach ($et in $eorzeanTimes) {
                if ($row.ContainsKey($et)) {
                    $rowData += $row[$et]
                } else {
                    $rowData += ""
                }
            }
            $dataTable.Rows.Add($rowData) | Out-Null
        }
    }

    # Initial population of the DataTable
    $rows = Group-DataIntoRows -data $data
    Populate-DataTable -rows $rows

    # Create the form
    $form = New-Object 'System.Windows.Forms.Form'
    $form.Text = "Eureka Weather Table"
    $form.Size = New-Object System.Drawing.Size(800, 420)
    $form.StartPosition = "CenterScreen"
    # Dark mode: Set form background color
    $form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $form.ShowIcon = $false
    # Hide the system title bar
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None

    # Create a custom title bar using a Panel
    $titleBar = New-Object System.Windows.Forms.Panel
    $titleBar.Dock = [System.Windows.Forms.DockStyle]::Top
    $titleBar.Height = 30
    $titleBar.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)


    
    # Create a PictureBox control
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Location = New-Object System.Drawing.Point(3,2)
    $pictureBox.Size = New-Object System.Drawing.Size(28, 28)
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom  # Adjust the image size to fit the PictureBox
    
    # Load the PNG image into the PictureBox
    $imagePath = "$PSScriptRoot\assets\orange.png"
    if (Test-Path $imagePath) {
        $pictureBox.Image = [System.Drawing.Image]::FromFile($imagePath)
    } else {
        Write-Host "Image not found!"
    }
    $titleBar.Controls.Add($pictureBox)

    # Add a label for the title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "     Eureka Weather Table"
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $titleBar.Controls.Add($titleLabel)

    # Add a close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "X"
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $closeButton.ForeColor = [System.Drawing.Color]::White
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.FlatAppearance.BorderSize = 0
    $closeButton.Dock = [System.Windows.Forms.DockStyle]::Right
    $closeButton.Width = 40
    $closeButton.Add_Click({
        $form.Close()
    })
    $titleBar.Controls.Add($closeButton)

    # Add the custom title bar to the form
    $form.Controls.Add($titleBar)

    # Define global variables for dragging
    $global:dragging = $false
    $global:dragStartPoint = New-Object System.Drawing.Point

    # MouseDown Event: Capture the starting point when the user clicks
    $titleLabel.Add_MouseDown({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $global:dragging = $true
            $global:dragStartPoint = $e.Location
        }
    })

    # MouseMove Event: Move the form if dragging is active
    $titleLabel.Add_MouseMove({
        param($sender, $e)
        if ($global:dragging) {
            $newLocation = $form.PointToScreen($e.Location)
            $form.Location = New-Object System.Drawing.Point(($newLocation.X - $global:dragStartPoint.X), ($newLocation.Y - $global:dragStartPoint.Y))
        }
    })

    # MouseUp Event: Stop dragging when the mouse button is released
    $titleLabel.Add_MouseUp({
        param($sender, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $global:dragging = $false
        }
    })

    # Create a Panel to hold the DataGridView
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Location = New-Object System.Drawing.Point(0, $titleBar.Height) # Start below the title bar
    $contentPanel.Size = New-Object System.Drawing.Size($form.Size.Width,$form.ClientSize.Height)
    $form.Controls.Add($contentPanel)


    # Create the DataGridView and bind it to the DataTable
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.DataSource = $dataTable
    $dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $dataGridView.AutoResizeColumnHeadersHeight()
    $dataGridView.RowTemplate.Height = 28
    #$dataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
    $dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dataGridView.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dataGridView.ScrollBars = [System.Windows.Forms.ScrollBars]::None

    # Dark mode: Customize DataGridView appearance
    $dataGridView.BackgroundColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $dataGridView.GridColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $dataGridView.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $dataGridView.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $dataGridView.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $dataGridView.EnableHeadersVisualStyles = $false
    $dataGridView.RowHeadersVisible = $false

    # Disable cell selection
    $dataGridView.ReadOnly = $true  # Make the DataGridView read-only
    $dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dataGridView.ClearSelection()  # Clear any initial selection

    # Handle the SelectionChanged event to prevent selection
    $dataGridView.Add_SelectionChanged({
        $dataGridView.ClearSelection()
    })

    # Highlight cells where jTime matches highlightedTime
    # Handle the CellFormatting event to highlight cells
    # Highlight cells where jTime matches highlightedTime
    $dataGridView.Add_CellFormatting({
        param($sender, $e)
        $cellValue = $e.Value
        if ($cellValue -isnot [System.DBNull]) {
            if ($cellValue -and $cellValue.StartsWith($highlightedTime)) {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(173, 216, 230)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
            } else {
                $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
                $e.CellStyle.ForeColor = [System.Drawing.Color]::White
            }
        }
    })

    function Increase-DataGridViewFontSize {
        param (
            [System.Windows.Forms.DataGridView]$dgv,
            [float]$increaseBy
        )
    
        # Get the current font of the DataGridView
        $currentFont = $dgv.Font
    
        # Create a new font with the increased size
        $newFont = New-Object System.Drawing.Font($currentFont.FontFamily, ($currentFont.Size + $increaseBy), $currentFont.Style)
    
        # Apply the new font to the DataGridView
        $dgv.Font = $newFont
    }

    # Increase the font size of the datagridview
    Increase-DataGridViewFontSize -dgv $dataGridView -increaseBy 3

    # Add the DataGridView to the content panel
    $contentPanel.Controls.Add($dataGridView)

    # Create labels for displaying information
    $timeLabel = New-Object System.Windows.Forms.Label
    $timeLabel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $timeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $timeLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $timeLabel.ForeColor = [System.Drawing.Color]::White  # White text
    $timeLabel.Height = 30
    $timeLabel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

    $specialLabel = New-Object System.Windows.Forms.Label
    $specialLabel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $specialLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $specialLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $specialLabel.ForeColor = [System.Drawing.Color]::White  # White text
    $specialLabel.Height = 30
    $specialLabel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

    $zoneLabel = New-Object System.Windows.Forms.Label
    $zoneLabel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $zoneLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $zoneLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $zoneLabel.ForeColor = [System.Drawing.Color]::White  # White text
    $zoneLabel.Height = 30
    $zoneLabel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)

    $currentWeather = $data | Where-Object {
        $_.jTime.StartsWith($highlightedTime)
    } | Select-Object -First 1
    $zoneLabel.Text = "Zone: $zone | Current Weather: $($currentWeather.jWeather)"

    # Function to update the labels
    function Update-Labels {
        $currentTime = Get-Date

        # Find the next jTime and its weather
        $nextJTime = $data | ForEach-Object {
            $time = [datetime]::ParseExact($_.jTime.Replace(" PM", "").Replace(" AM", ""), "HH:mm", $null)
            if ($time.TimeOfDay -lt $currentTime.TimeOfDay) {
                $time = $time.AddDays(1) # Wrap around to the next day
            }
            [PSCustomObject]@{
                Time = $time
                Weather = $_.jWeather
            }
        } | Sort-Object Time | Select-Object -First 1

        if ($nextJTime) {
            $timeDifference = $nextJTime.Time - $currentTime
            $minutes = [math]::Floor($timeDifference.TotalMinutes)
            $seconds = $timeDifference.Seconds
            $timeLabel.Text = "Current Time: $($currentTime.ToString("HH:mm tt")) | Time until $($nextJTime.Weather): $(("{0:D2}m {1:D2}s" -f [string]$minutes, [string]$seconds))"
        } else {
            $timeLabel.Text = "Current Time: $($currentTime.ToString("HH:mm tt")) | No upcoming jTime"
        }

        # Special label logic for Blizzard and Fog (only if $zone -eq "Pagos")
        if ($zone -eq "Pagos") {
            $nextBlizzard = $data | ForEach-Object {
                if ($_.jWeather -eq "Blizzard") {
                    $time = [datetime]::ParseExact($_.jTime.Replace(" PM", "").Replace(" AM", ""), "HH:mm", $null)
                    if ($time.TimeOfDay -lt $currentTime.TimeOfDay) {
                        $time = $time.AddDays(1) # Wrap around to the next day
                    }
                    [PSCustomObject]@{
                        Time = $time
                    }
                }
            } | Sort-Object Time | Select-Object -First 1

            $nextFog = $data | ForEach-Object {
                if ($_.jWeather -eq "Fog") {
                    $time = [datetime]::ParseExact($_.jTime.Replace(" PM", "").Replace(" AM", ""), "HH:mm", $null)
                    if ($time.TimeOfDay -lt $currentTime.TimeOfDay) {
                        $time = $time.AddDays(1) # Wrap around to the next day
                    }
                    [PSCustomObject]@{
                        Time = $time
                    }
                }
            } | Sort-Object Time | Select-Object -First 1

            $blizzardText = if ($nextBlizzard) {
                $blizzardDiff = $nextBlizzard.Time - $currentTime
                $blizzardMinutes = [math]::Floor($blizzardDiff.TotalMinutes)
                $blizzardSeconds = $blizzardDiff.Seconds
                "Blizzard: $(("{0:D2}m {1:D2}s" -f [string]$blizzardMinutes, [string]$blizzardSeconds))"
                if (($blizzardMinutes -eq 15 -or $blizzardMinutes -eq 10) -and ($blizzardSeconds -eq 0)){
                    WeatherAlarm
                }
            } else {
                "Blizzard: None"
            }

            $fogText = if ($nextFog) {
                $fogDiff = $nextFog.Time - $currentTime
                $fogMinutes = [math]::Floor($fogDiff.TotalMinutes)
                $fogSeconds = $fogDiff.Seconds
                "Fog: $(("{0:D2}m {1:D2}s" -f [string]$fogMinutes, [string]$fogSeconds))"
                if (($fogMinutes -eq 15 -or $fogMinutes -eq 10) -and ($fogSeconds -eq 0)){
                    WeatherAlarm
                }
            } else {
                "Fog: None"
            }

            $specialLabel.Text = "$blizzardText | $fogText"
        } else {
            $specialLabel.Text = ""
        }

        if (([datetime]::ParseExact(($highlightedTime).Replace(" PM", "").Replace(" AM", ""), "HH:mm", $null).AddSeconds(1400)) -lt ($currentTime)) {
            $highlightedTime = (GetIntervalStart).ToString("HH:mm")
            $currentWeather = $data | Where-Object {
                $_.jTime.StartsWith($highlightedTime)
            } | Select-Object -First 1
            
            $zoneLabel.Text = "Zone: $zone | Current Weather: $($currentWeather.jWeather)"
            # Fetch new JSON data
            $jsonData = GetWeatherDataJson -zone $zone
            $data = $jsonData | ConvertFrom-Json

            # Regroup data into rows
            $rows = Group-DataIntoRows -data $data

            # Repopulate the DataTable
            Populate-DataTable -rows $rows

            #Force the DataGridView to refresh
            $dataGridView.Refresh()
        }
        # Determine the current weather based on the highlighted time (Cyan cell)

    }

    # Timer to update the labels every second
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000 # 1 second
    $timer.Add_Tick({
        Update-Labels
    })
    $timer.Start()


    # Add the labels to the form
    $contentPanel.Controls.Add($zoneLabel)
    $contentPanel.Controls.Add($timeLabel)
    $contentPanel.Controls.Add($specialLabel)

    # Show the form
    $form.ShowDialog()
}

DisplayWeatherTable -zone Pagos