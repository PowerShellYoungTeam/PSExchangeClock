Function Start-Countdown{

$script:Canceling = $False

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$DueDate = Get-Date -Date "24/12/2021" -Hour 12 -Minute 30 #Due Date/Time!

$Form  = New-Object system.Windows.Forms.Form
$Form.Width   = "300"
$Form.Height  = "200"
$Form.Text = "LSE Close"
$Form.BackColor  = "Black"
$lblCountDown = New-Object System.Windows.Forms.Label
$lblCountDown.AutoSize = $true
$lblCountDown.Font = "Impact, 20"
$lblCountDown.ForeColor = "Green"
$lblCountDown.BackColor = "Black"
$lblCountDown.TextAlign = "MiddleCenter"
$lblCountDown.Text =    
  "LSE closes in: {0:hh}:{0:mm}:{0:ss}" -f 
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
    "LSE closes in: {0:hh}:{0:mm}:{0:ss}" -f 
     ($DueDate - (Get-Date))
 }

)

# Progress bar:

$tmrCountdown.Start()
$Form.ShowDialog()
$tmrCountdown.Dispose()
Remove-Variable tmrCountdown
}


Write-Host "Started: $(Get-date)"

$Dropout = Start-Countdown

if($True -eq $script:Canceling){
    Exit
    }

Write-host ""
Write-host -ForegroundColor Green "###########################################"
Write-host -ForegroundColor Green "#                                         #"
Write-host -ForegroundColor Green "#             MARKETS CLOSED              #"
Write-host -ForegroundColor Green "#                                         #"
Write-host -ForegroundColor Green "###########################################"
Write-host ""
pause
