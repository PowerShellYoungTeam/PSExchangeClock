Function Start-Countdown{

$script:Canceling = $False

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$DueDate = Get-Date -Date "24/12/2021" -Hour 16 -Minute 00 #Due Date/Time!

$Form  = New-Object system.Windows.Forms.Form
$Form.Width   = "400"
$Form.Height  = "200"
$Form.Text = "Time to Pop Smoke and bug out!"
$Form.BackColor  = "Black"
$lblCountDown = New-Object System.Windows.Forms.Label
$lblCountDown.AutoSize = $true
$lblCountDown.Font = "Impact, 20"
$lblCountDown.ForeColor = "Green"
$lblCountDown.BackColor = "Black"
$lblCountDown.TextAlign = "MiddleCenter"
$lblCountDown.Text =    
  "Popping Smoke in: {0:hh}:{0:mm}:{0:ss}" -f 
  ($DueDate - (Get-Date))

$Form.Controls.Add($lblCountDown)

$cancellFont = New-Object System.Drawing.Font('Impact',25,[System.Drawing.FontStyle]::Bold)
$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.AutoSize = $true 
$CloseButton.Text = "CANCEL"
$CloseButton.ForeColor = "BLACK"
$CloseButton.BackColor = "RED"
$CloseButton.Font = $cancellFont
$CloseButton.Top = 40
$CloseButton.Add_Click({Write-Host "Cancelled at - $(get-date)" -Fore ReD
			$script:Canceling=$true
			[System.Windows.Forms.Application]::DoEvents()
			$Form.Close()})
$form.CancelButton = $CloseButton
$Form.Controls.Add($CloseButton)
$tmrCountdown = New-Object System.Windows.Forms.Timer
$tmrCountdown.Interval = 1000
$tmrCountdown.add_Tick(
 {
    If($DueDate.AddMinutes(-30) -le (get-date)){
        $lblCountDown.ForeColor = "Yellow"
    }
     If($DueDate.AddMinutes(-15) -le (get-date)){
        $lblCountDown.ForeColor = "Red"
    }

      If($DueDate -le (Get-Date)){
        $Form.Close()
    }
  $lblCountDown.Text = 
    "Popping Smoke in: {0:hh}:{0:mm}:{0:ss}" -f 
     ($DueDate - (Get-Date))
 }

)

# Progress bar:

$tmrCountdown.Start()
$Form.ShowDialog()
$tmrCountdown.Dispose()
Remove-Variable tmrCountdown
}

# Start script
Write-Host "Started: $(Get-date)"

$Dropout = Start-Countdown

if($True -eq $script:Canceling){
    Exit
    }

Write-host ""
Write-host -ForegroundColor Green "###########################################"
Write-host -ForegroundColor Green "#                                         #"
Write-host -ForegroundColor Green "#             Popping Smoke               #"
Write-host -ForegroundColor Green "#                                         #"
Write-host -ForegroundColor Green "###########################################"
Write-host ""

#close down apps
$Apps = ("Outlook",
         "mstsc",
         "Powershell_Ise",
         "Teams",
         "Chrome",
         "msedge")

#Kill processes
foreach($App in $Apps){
    Try{
        get-Process -Name $App -ErrorAction Stop| Stop-Process -Force -ErrorAction Stop 
        Write-host -ForegroundColor Yellow "$($App) Closed"
    }Catch{
        $ErrorMessage = $_.Exception.Message
        Write-host -ForegroundColor RED "Issue Killing  $($App) because $($ErrorMessage)"                  
    }
}# end of foreach

Write-host ""
Write-host -ForegroundColor Green "###########################################"
Write-host -ForegroundColor Green "#                                         #"
Write-host -ForegroundColor Green "#             Bugging Out                 #"
Write-host -ForegroundColor Green "#                                         #"
Write-host -ForegroundColor Green "###########################################"
Write-host ""

# SHutdown machine

Stop-Computer -ComputerName localhost -Force 

pause
