function Log-FormData {
    param (
        [string]$ClientName,
        [string]$TicketNumber,
        [string]$Description,
        [string]$Pin,
        [int]$MaintenanceDuration,
        [switch]$DateInFile
    )
    # Determine if date should be included in the filename
    $date = if ($DateInFile) { "-$(Get-Date -Format 'yyyy-MM-dd')" } else { "" }
    $yearMonth = Get-Date -Format "yyyy-MM"
    $today = Get-Date -Format "yyyy-MM-dd"

    # Define log file paths with optional date in the filename
    $logFilePath1 = "\\FILE_SHARE\Logs\LoginAudit\Users\$env:username\$env:username$date.csv"
    $logFilePath2 = "\\FILE_SHARE\Logs\LoginAudit\Servers\$env:computername\$env:computername$date.csv"
    $logFilePath3 = "\\FILE_SHARE\Logs\LoginAudit\Global\LoginAuditFull$date.csv"
    $LoginAuditMonth = "\\FILE_SHARE\Logs\LoginAudit\Global\LoginAuditMonth-$yearMonth$date.csv"
    $dailyLogFilePath = "\\FILE_SHARE\Logs\LoginAudit\Daily\LoginAudit-$today.csv"

    # Get the directory paths
    $dir1 = [System.IO.Path]::GetDirectoryName($logFilePath1)
    $dir2 = [System.IO.Path]::GetDirectoryName($logFilePath2)
    $dir3 = [System.IO.Path]::GetDirectoryName($logFilePath3)
    $dir4 = [System.IO.Path]::GetDirectoryName($LoginAuditMonth)
    $dir5 = [System.IO.Path]::GetDirectoryName($dailyLogFilePath)

    # Create directories if they do not exist
    if (-not (Test-Path -Path $dir1)) {
        New-Item -Path $dir1 -ItemType Directory -Force
    }
    if (-not (Test-Path -Path $dir2)) {
        New-Item -Path $dir2 -ItemType Directory -Force
    }
    if (-not (Test-Path -Path $dir3)) {
        New-Item -Path $dir3 -ItemType Directory -Force
    }
    if (-not (Test-Path -Path $dir4)) {
        New-Item -Path $dir4 -ItemType Directory -Force
    }
    if (-not (Test-Path -Path $dir5)) {
        New-Item -Path $dir5 -ItemType Directory -Force
    }

    # Create a single-line log entry as a custom object
    $logEntry = [PSCustomObject]@{
        Date = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Server = $env:computername
        Username = $env:username
        ClientName = $ClientName
        TicketNumber = $TicketNumber
        Description = $Description
        Pin = $Pin
        MaintenanceDuration = $MaintenanceDuration
    }

    # Append the log entry to all relevant CSV files
    $csvExists1 = Test-Path -Path $logFilePath1
    $csvExists2 = Test-Path -Path $logFilePath2
    $csvExists3 = Test-Path -Path $logFilePath3
    $csvExists4 = Test-Path -Path $LoginAuditMonth
    $dailyLogExists = Test-Path -Path $dailyLogFilePath

    if ($csvExists1) {
        $logEntry | Export-Csv -Path $logFilePath1 -Append -NoTypeInformation
    } else {
        $logEntry | Export-Csv -Path $logFilePath1 -NoTypeInformation
    }
    if ($csvExists2) {
        $logEntry | Export-Csv -Path $logFilePath2 -Append -NoTypeInformation
    } else {
        $logEntry | Export-Csv -Path $logFilePath2 -NoTypeInformation
    }
    if ($csvExists3) {
        $logEntry | Export-Csv -Path $logFilePath3 -Append -NoTypeInformation
    } else {
        $logEntry | Export-Csv -Path $logFilePath3 -NoTypeInformation
    }
    if ($csvExists4) {
        $logEntry | Export-Csv -Path $LoginAuditMonth -Append -NoTypeInformation
    } else {
        $logEntry | Export-Csv -Path $LoginAuditMonth -NoTypeInformation
    }
    if ($dailyLogExists) {
        $logEntry | Export-Csv -Path $dailyLogFilePath -Append -NoTypeInformation
    } else {
        $logEntry | Export-Csv -Path $dailyLogFilePath -NoTypeInformation
    }
}



# Import required .NET namespaces
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
$azure = 0
$SendPINEmail = 1
# Define the XAML for the custom message box
$messageBoxXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Info" Height="100" Width="400" WindowStartupLocation="CenterScreen" WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Border Background="White" BorderBrush="Black" BorderThickness="1" CornerRadius="10" Padding="10">
        <TextBlock Name="MessageTextBlock" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="12"/>
    </Border>
</Window>
"@
function Get-BypassKey {
    param (
        [switch]$UseKey,
        [int]$KeyLength = 8,
        [string]$Username = $env:USERNAME,
        [string]$CustomKey,
        [switch]$ShowKey,
        [switch]$ShowBypassKey,
        [switch]$ShowKeyAndBypassKey,
        [switch]$UseSalt,
        [int]$SaltLength = 8
    )

    # Function to compute the MD5 hash
    function Get-MD5Hash {
        param (
            [Parameter(Mandatory=$true)]
            [string]$InputString
        )
        
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($InputString)
        $hash = $md5.ComputeHash($bytes)
        return [BitConverter]::ToString($hash) -replace '-'
    }

    # Function to generate a random key or salt
    function Get-RandomString {
        param (
            [Parameter(Mandatory=$true)]
            [int]$Length
        )
        
        $chars = 48..57 + 65..90 + 97..122 | ForEach-Object {[char]$_}
        -join (1..$Length | ForEach-Object { $chars | Get-Random })
    }

    # Get the current date in a specific format (e.g., yyyy-MM-dd)
    $date = Get-Date -Format "yyyy-MM-dd"

    # Prepare the output object
    $output = [PSCustomObject]@{
        Key = $null
        BypassKey = $null
        Salt = $null
    }

    # Determine the input string and output based on parameters
    if ($ShowKeyAndBypassKey) {
        if ($UseKey) {
            if ($CustomKey) {
                $key = $CustomKey
            } else {
                $key = Get-RandomString -Length $KeyLength
            }
            $output.Key = $key
            $inputString = "$Username$date$key"
        } else {
            $output.Key = "No key is used."
            $inputString = "$Username$date"
        }
        
        if ($UseSalt) {
            $salt = Get-RandomString -Length $SaltLength
            $output.Salt = $salt
            $inputString = "$inputString$salt"
        }

        # Get the MD5 hash of the input string
        $md5Hash = Get-MD5Hash -InputString $inputString
        $output.BypassKey = $md5Hash
        Write-Output $output
        return
    }

    if ($ShowBypassKey) {
        if ($UseKey) {
            if ($CustomKey) {
                $key = $CustomKey
            } else {
                $key = Get-RandomString -Length $KeyLength
            }
            $inputString = "$Username$date$key"
        } else {
            $inputString = "$Username$date"
        }
        
        if ($UseSalt) {
            $salt = Get-RandomString -Length $SaltLength
            $output.Salt = $salt
            $inputString = "$inputString$salt"
        }

        # Get the MD5 hash of the input string and output only the hash
        $md5Hash = Get-MD5Hash -InputString $inputString
        $output.BypassKey = $md5Hash
        Write-Output $output.BypassKey
        return
    }

    if ($ShowKey) {
        if ($UseKey) {
            if ($CustomKey) {
                $key = $CustomKey
            } else {
                $key = Get-RandomString -Length $KeyLength
            }
            $output.Key = $key
        } else {
            $output.Key = "No key is used."
        }
        Write-Output $output.Key
        return
    }

    # Determine the key and input string for regular output
    if ($UseKey) {
        if ($CustomKey) {
            $key = $CustomKey
        } else {
            $key = Get-RandomString -Length $KeyLength
        }
        $output.Key = $key
        $inputString = "$Username$date$key"
    } else {
        $inputString = "$Username$date"
    }

    if ($UseSalt) {
        $salt = Get-RandomString -Length $SaltLength
        $output.Salt = $salt
        $inputString = "$inputString$salt"
    }

    # Get the MD5 hash of the input string
    $md5Hash = Get-MD5Hash -InputString $inputString
    $output.BypassKey = $md5Hash

    # Output the result
    Write-Output "Username: $Username"
    Write-Output "Date: $date"
    Write-Output "Key: $($output.Key)"
    Write-Output "Salt: $($output.Salt)"
    Write-Output "BypassKey: $($output.BypassKey)"
}

$bypasspingkeytoday = Get-BypassKey -usekey -ShowKeyAndBypassKey
$bypasspingkeytoday1 = $bypasspingkeytoday.key
#$bypasspingkeytoday1
$pin2use = $bypasspingkeytoday.BypassKey
#$pin2use

# Function to show the custom message box and auto-close it
function Show-AutoCloseMessageBox {
    param (
        [string]$message
    )

    $xmlReader = New-Object System.Xml.XmlTextReader -ArgumentList (New-Object System.IO.StringReader $messageBoxXaml)
    $messageBoxWindow = [Windows.Markup.XamlReader]::Load($xmlReader)

    $messageTextBlock = $messageBoxWindow.FindName("MessageTextBlock")
    $messageTextBlock.Text = $message

    # Timer to close the window after 5 seconds
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000  # 5000 milliseconds = 5 seconds
    $timer.Add_Tick({
        $messageBoxWindow.Close()
        $timer.Stop()
    })
    $timer.Start()

    $messageBoxWindow.ShowDialog()
}

# Function to send an email with the PIN
function Send-PINEmail {
    param (
        [string]$emailAddress,
        [string]$pin
        #[string]$bypasspingkeytoday1
    )
    # Check if $SendPINEmail is set to 0, if so, do nothing
    if ($SendPINEmail -eq 0) {
    write-host Email disabled
        return
    }
    $smtpFrom = "dc@email.com"
    $messageSubject = "Your DC Login PIN"
    $messageBody = "Your Login PIN is: $pin `nKey: $bypasspingkeytoday1 `nConnected From $env:CLIENTNAME `nAccount $env:USERNAME `nServer $env:COMPUTERNAME `nReason $description Ref:$clientName,$ticketNumber,By-$env:username (global) $emailAddress"

    $mailMessage = New-Object system.net.mail.mailmessage
    $mailMessage.From = $smtpFrom
    $mailMessage.To.Add($emailAddress)
    $mailMessage.Bcc.Add("dcnotice@email.com")  # Add BCC recipient
    $mailMessage.Subject = $messageSubject
    $mailMessage.Body = $messageBody

    $smtpServers = @("10.5.7.90", "10.5.7.91", "10.5.7.92")
    $smtpClient = $null
    $connected = $false

    foreach ($smtpServer in $smtpServers) {
        # Test connection to SMTP server on port 25 for 1 second
        try {
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $asyncResult = $tcpClient.BeginConnect($smtpServer, 25, $null, $null)
            $waitHandle = $asyncResult.AsyncWaitHandle
            if ($waitHandle.WaitOne(1000, $false)) {
                $tcpClient.EndConnect($asyncResult)
                $tcpClient.Close()
                Write-Host "Connection to SMTP server $smtpServer on port 25 successful."
                $connected = $true
                $smtpClient = New-Object Net.Mail.SmtpClient($smtpServer)
                break
            } else {
                $tcpClient.Close()
                throw [System.TimeoutException] "Connection to SMTP server $smtpServer timed out."
            }
        } catch {
            Write-Host "Failed to connect to SMTP server $smtpServer on port 25. Error: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($connected) {
        try {
            $ErrorActionPreference = "Stop"
            $smtpClient.Send($mailMessage)
            Write-Host "Email sent successfully to $emailAddress."
        } catch {
            Write-Host "Login Audit failed to send the PIN to $emailAddress. Error: $($_.Exception.Message)" -ForegroundColor Yellow
            if ($_.Exception.InnerException) {
                Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
            }
        } finally {
            $ErrorActionPreference = "Continue"
        }
    } else {
        $message = "Login Audit failed to send the PIN to $emailAddress. Please use your ByPassPIN"
        Write-Host $message -ForegroundColor Red
       # [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        [System.Windows.Forms.MessageBox]::Show($message, "SMTP Failure", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}



function Send-LoginEmail {
    param (
        [string]$emailAddress,
        [string]$pin
    )
    # Check if $SendPINEmail is set to 0, if so, do nothing
    if ($SendPINEmail -eq 0) {
    write-host Email disabled
        return
    }
    $smtpServer = "10.5.7.90"
    $smtpFrom = "dc@email.com"
    $messageSubject = "Login Audit Logon Passed"
  # $messageBody = "Connected From $env:CLIENTNAME `nAccount $env:USERNAME `nServer $env:computername"
$messageBody = "Your Login PIN is: $pin `nKey: $bypasspingkeytoday1 `nConnected From $env:CLIENTNAME `nAccount $env:USERNAME `nServer $env:computername `nReason $description Ref:$clientName,$ticketNumber,By-$env:username (global) $emailAddress"


    $mailMessage = New-Object system.net.mail.mailmessage
    $mailMessage.From = $smtpFrom
    $mailMessage.To.Add("")
    $mailMessage.Subject = $messageSubject
    $mailMessage.Body = $messageBody

    $smtpClient = New-Object Net.Mail.SmtpClient($smtpServer)
    $smtpClient.Send($mailMessage)
}
# Function to create and write an event log entry
function Write-CustomEventLog {
    param (
        [string]$message
    )

    $logName = "Application"  # You can use a custom log name if desired
    $source = "LoginAudit"  # Event source name

    # Check if the event source exists
    if (-not [System.Diagnostics.EventLog]::SourceExists($source)) {
        # Create a new event source
        [System.Diagnostics.EventLog]::CreateEventSource($source, $logName)
    }

    # Write the event log entry with ID 100
    Write-EventLog -LogName $logName -Source $source -EventId 100 -EntryType Information -Message $message
}



# Update the Update-MaintenanceFile function to include event logging
function Update-MaintenanceFile {
    param (
        [string]$description,
        [int]$maintenanceDuration,
        [bool]$extended,
        [string]$clientName,
        [string]$ticketNumber
    )

    # Check if $azure is set to 1, if so, do nothing
    if ($azure -eq 1) {
        return
    }
    
    # Do not update the file if maintenance duration is 0
    if ($maintenanceDuration -le 0) {
        return
    }

    $serverName = "$env:COMPUTERNAME"
    $timeL = $maintenanceDuration
    $reasonL = "ApplicationInstallation"
    $nameL = "$env:USERNAME"
    $filePath = "\\backup02\cs\Blue_Install\MaintenanceMode\MaintenanceMode.txt"

    # Format commentL
    $commentL = "$description Ref:$clientName,$ticketNumber,By-$nameL"
    if ($extended) {
        $commentL += " (extended)"
    }

    $content = @"
servername=$serverName
TimeL=$timeL
commentL=$commentL
nameL=$nameL
reasonL=$reasonL
"@

    $content | Set-Content -Path $filePath

    # Write the event log entry with the same info as commentL
    #Write-CustomEventLog -message $commentL
}


# Initialize countdown window and timer variables
$countdownWindow = $null
$countdownTimer = $null
$remainingTime = 0

function Initialize-CountdownWindow {
    param (
        [int]$duration
    )

    $global:remainingTime = [math]::Max($duration, 6) * 60  # Ensure minimum duration is 6 minutes

    # Define the XAML for the countdown window
    $countdownXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Maintenance Timer"

        Height="200" Width="300" WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Name="CountdownLabel" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="36"/>
        <Button Content="Extend" Name="ExtendButton" HorizontalAlignment="Left" VerticalAlignment="Bottom" Width="75" Margin="10"/>
        <Button Content="Cancel" Name="CancelButton" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="75" Margin="10"/>
    </Grid>
</Window>
"@

    $xmlReader = New-Object System.Xml.XmlTextReader -ArgumentList (New-Object System.IO.StringReader $countdownXaml)
    $global:countdownWindow = [Windows.Markup.XamlReader]::Load($xmlReader)

    # Set the window to always stay on top
    $global:countdownWindow.Topmost = $true

    # Define UI elements for the countdown window
    $countdownLabel = $global:countdownWindow.FindName("CountdownLabel")
    $extendButton = $global:countdownWindow.FindName("ExtendButton")
    $cancelButton = $global:countdownWindow.FindName("CancelButton")

    # Timer initialization
    $global:countdownTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:countdownTimer.Interval = [TimeSpan]::FromSeconds(1)
    $global:countdownTimer.Add_Tick({
        if ($global:remainingTime -gt 0) {
            $global:remainingTime -= 1
            $minutes = [math]::Floor($global:remainingTime / 60)
            $seconds = $global:remainingTime % 60
            $countdownLabel.Content = "{0}:{1:D2}" -f $minutes, $seconds
        } else {
            $global:countdownTimer.Stop()
            # Prompt to add more time in a new dialog
            $extendPromptXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Maintenance Ended"
        Height="150" Width="300" WindowStartupLocation="CenterScreen" WindowStyle="SingleBorderWindow">
    <Grid>
        <TextBlock Text="Maintenance period has ended. Would you like to add more time?" HorizontalAlignment="Center" VerticalAlignment="Center" TextWrapping="Wrap" Margin="10"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Bottom">
            <Button Content="Yes" Name="YesButton" Width="75" Margin="5"/>
            <Button Content="No" Name="NoButton" Width="75" Margin="5"/>
        </StackPanel>
    </Grid>
</Window>
"@

            $xmlReaderPrompt = New-Object System.Xml.XmlTextReader -ArgumentList (New-Object System.IO.StringReader $extendPromptXaml)
            $extendPromptWindow = [Windows.Markup.XamlReader]::Load($xmlReaderPrompt)
            $extendPromptWindow.Topmost = $true

            # Define UI elements for the extend prompt window
            $yesButton = $extendPromptWindow.FindName("YesButton")
            $noButton = $extendPromptWindow.FindName("NoButton")

            # Yes button event handler
            if ($yesButton) {
                $yesButton.Add_Click({
                    $global:remainingTime += 300  # Add 5 minutes (300 seconds) to remaining time
                    $minutes = [math]::Floor($global:remainingTime / 60)
                    $seconds = $global:remainingTime % 60
                    $countdownLabel.Content = "{0}:{1:D2}" -f $minutes, $seconds

                    # Update the maintenance file with the extended description
                    $description = $descriptionTextBox.Text
                    Update-MaintenanceFile -description $description -maintenanceDuration ($global:remainingTime / 60) -extended $true

                    # Close the extend prompt and restart the countdown
                    $extendPromptWindow.Close()
                    $global:countdownTimer.Start()
                })
            }

            # No button event handler
            if ($noButton) {
                $noButton.Add_Click({
                    # Close the extend prompt and end maintenance
                    $extendPromptWindow.Close()
                    Close-Windows
                })
            }

            # Show the extend prompt window
            $extendPromptWindow.ShowDialog()
        }
    })
    $global:countdownTimer.Start()

    # Extend button event handler
    if ($extendButton) {
        $extendButton.Add_Click({
            $global:remainingTime += 300  # Add 5 minutes (300 seconds) to remaining time
            $minutes = [math]::Floor($global:remainingTime / 60)
            $seconds = $global:remainingTime % 60
            $countdownLabel.Content = "{0}:{1:D2}" -f $minutes, $seconds

            # Update the maintenance file with the extended description
            $description = $descriptionTextBox.Text
            Update-MaintenanceFile -description $description -maintenanceDuration ($global:remainingTime / 60) -extended $true
        })
    }

    # Cancel button event handler
    if ($cancelButton) {
        $cancelButton.Add_Click({
            if ($global:countdownTimer) {
                $global:countdownTimer.Stop()
            }
            Close-Windows

        })
    }

    # Minimize the main window and show the countdown window
    $window.WindowState = 'Minimized'
    $global:countdownWindow.ShowDialog()
}


# Function to close all windows
function Close-Windows {
    if ($global:countdownWindow) {
        $global:countdownWindow.Close()
    }
    # Restore the main window if minimized
    if ($window.WindowState -eq 'Minimized') {
        $window.WindowState = 'Normal'
    }
    # Close the main window
    $window.Close()
}

# Function to run explorer.exe
function Run-Explorer {
Send-LoginEmail
    Start-Process "userinit.exe"
}

# Define the XAML for the main window
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Login Audit"
WindowStyle="None" 
        Height="400" Width="400"
	WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Label Grid.Row="0" Grid.Column="0" Content="Client Name:" VerticalAlignment="Center"/>
        <TextBox Grid.Row="0" Grid.Column="1" Name="ClientNameTextBox" Margin="5"/>

        <Label Grid.Row="1" Grid.Column="0" Content="Ticket Number:" VerticalAlignment="Center"/>
        <TextBox Grid.Row="1" Grid.Column="1" Name="TicketNumberTextBox" Margin="5"/>

        <Label Grid.Row="2" Grid.Column="0" Content="Description:" VerticalAlignment="Center"/>
        <TextBox Grid.Row="2" Grid.Column="1" Name="DescriptionTextBox" Margin="5" AcceptsReturn="True" Height="60"/>

        <Label Grid.Row="3" Grid.Column="0" Content="PIN:" VerticalAlignment="Center"/>
        <PasswordBox Grid.Row="3" Grid.Column="1" Name="PinTextBox" Margin="5"/>

        <Label Grid.Row="4" Grid.Column="0" Content="Maintenance Duration (mins):" VerticalAlignment="Center"/>
        <TextBox Grid.Row="4" Grid.Column="1" Name="MaintenanceDurationTextBox" Margin="5" Text="0"/>

        <Button Grid.Row="5" Grid.Column="0" Name="SubmitButton" Content="Submit" Margin="5"/>
        <Button Grid.Row="5" Grid.Column="1" Name="ClearButton" Content="Clear" Margin="5" HorizontalAlignment="Right"/>
        
        <Grid Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button Grid.Column="0" Name="CancelButton" Content="Cancel" Margin="5" VerticalAlignment="Bottom"/>
            <StackPanel Grid.Column="1" VerticalAlignment="Bottom">
                <TextBlock Text="Please check your email for your pin." HorizontalAlignment="Center" Margin="0,0,0,10"/>
		<TextBlock Text="The email came from dc@email.com." HorizontalAlignment="Center" Margin="0,0,0,10"/>
		<TextBlock Text="Contact Sean Joseph on Slack for immediate help" HorizontalAlignment="Center" Margin="0,0,0,10"/>
<TextBox Text="User $env:username" Margin="0,0,0,5" HorizontalAlignment="Right" IsReadOnly="True" BorderThickness="0" Background="Transparent"/>
<TextBox Text="Key: $bypasspingkeytoday1" Margin="0,0,0,5" HorizontalAlignment="Right" IsReadOnly="True" BorderThickness="0" Background="Transparent"/>
                <Image Name="LogoImage" Width="100" Height="50" HorizontalAlignment="Right" Margin="0,0,5,5"/>

            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@

# Load the main window XAML
$reader = New-Object System.Xml.XmlTextReader -ArgumentList (New-Object System.IO.StringReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls from the main window
$submit = $window.FindName("SubmitButton")
$clear = $window.FindName("ClearButton")
$cancel = $window.FindName("CancelButton")
$descriptionTextBox = $window.FindName("DescriptionTextBox")
$clientNameTextBox = $window.FindName("ClientNameTextBox")
$ticketNumberTextBox = $window.FindName("TicketNumberTextBox")
$pinTextBox = $window.FindName("PinTextBox")
$maintenanceDurationTextBox = $window.FindName("MaintenanceDurationTextBox")
$logoImage = $window.FindName("LogoImage")
$logoRow = $window.FindName("LogoRow")

# Check if the logo file exists
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logoPath = [System.IO.Path]::Combine($scriptDir, 'logo.jpg')

if (Test-Path $logoPath) {
    # Load the logo image
    $logoBitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $logoBitmap.BeginInit()
    $logoBitmap.UriSource = [System.Uri]::new($logoPath, [System.UriKind]::Absolute)
    $logoBitmap.EndInit()
    $logoImage.Source = $logoBitmap
    $logoImage.Visibility = 'Visible'
} else {
    # Hide logo
    $logoImage.Visibility = 'Collapsed'
}

# Function to generate a random PIN
function Generate-RandomPIN {
    return -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 4 | % {[char]$_})
}

# Function to get the user's email address from Active Directory
function Get-UserEmail {
    param (
        [string]$username
    )

    $user = ((([adsisearcher]"(&(objectCategory=User)(sAMAccountName=$env:USERNAME))").findall()).properties).mail
    if ([string]::IsNullOrEmpty($user)) {
        return $null
    }
    return $user
}

# Function to show a message box
function Show-MessageBox {
    param (
        [string]$message,
        [string]$title
    )

    [System.Windows.MessageBox]::Show($message, $title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
}

# Send the initial email with the PIN
$pin = Generate-RandomPIN
$emailAddress = Get-UserEmail $env:USERNAME

if ($SendPINEmail -ne 0) {
    if ($emailAddress) {
        Send-PINEmail -emailAddress $emailAddress -pin $pin
        Show-AutoCloseMessageBox -message "An email with the PIN has been sent to $emailAddress."
    } else {
        Show-MessageBox -message "No email address was found for your account. Please contact support." -title "Email Not Found"
    }
}

# Button click evgfent handlers
$submit.Add_Click({
    $description = $descriptionTextBox.Text
    $clientName = $clientNameTextBox.Text
    $ticketNumber = $ticketNumberTextBox.Text
    $enteredPin = $pinTextBox.Password
    $maintenanceDuration = $maintenanceDurationTextBox.Text -as [int]

   # Log form data
#    Log-FormData -ClientName $clientName -TicketNumber $ticketNumber -Description $description -Pin $pin -MaintenanceDuration $maintenanceDuration


    if (-not $clientName) { 
        [System.Windows.MessageBox]::Show("Client Name is required.") 
# Log form data
$clientName = "Client Name is required."
    Log-FormData -ClientName $clientName -TicketNumber $ticketNumber -Description $description -Pin $pin -MaintenanceDuration $maintenanceDuration

        return 
    }
    if (-not $description) { 
        [System.Windows.MessageBox]::Show("Description is required.") 
# Log form data
$description = "Description is required."
    Log-FormData -ClientName $clientName -TicketNumber $ticketNumber -Description $description -Pin $pin -MaintenanceDuration $maintenanceDuration

        return 
    }

    # Check if the maintenance duration is between 1 and 5 minutes
    if ($maintenanceDuration -lt 6 -and $maintenanceDuration -gt 0) {
        $maintenanceDuration = 6
    }

$emessage = "$description Ref:$clientName,$ticketNumber,By-$env:username (global) $emailAddress"
Write-CustomEventLog -message $emessage
    if ($enteredPin -eq "$pin2use") {
        [System.Windows.MessageBox]::Show("Bypass PIN used.", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        # Bypass PIN used, proceed without checking the generated PIN
        if ($maintenanceDuration -ge 6) {
		$description = $description + " BYPASSPIN Used"
            # Run explorer.exe before initializing the countdown
# Log form data
    Log-FormData -ClientName $clientName -TicketNumber $ticketNumber -Description $description -Pin $pin -MaintenanceDuration $maintenanceDuration

            Run-Explorer
            # Update maintenance file and start countdown
            Update-MaintenanceFile -description $description -maintenanceDuration $maintenanceDuration -extended $false -clientName $clientName -ticketNumber $ticketNumber
            Initialize-CountdownWindow -duration $maintenanceDuration
        } else {
            # No maintenance, just start explorer.exe
 # Log form data
    Log-FormData -ClientName $clientName -TicketNumber $ticketNumber -Description $description -Pin $pin -MaintenanceDuration $maintenanceDuration

            Run-Explorer
            Close-Windows }
    } elseif ($enteredPin -eq $pin) {
        if ($maintenanceDuration -ge 6) {
            # Run explorer.exe before initializing the countdown
 # Log form data
    Log-FormData -ClientName $clientName -TicketNumber $ticketNumber -Description $description -Pin $pin -MaintenanceDuration $maintenanceDuration

            Run-Explorer
            # Update maintenance file and start countdown
            Update-MaintenanceFile -description $description -maintenanceDuration $maintenanceDuration -extended $false -clientName $clientName -ticketNumber $ticketNumber
            Initialize-CountdownWindow -duration $maintenanceDuration
        } else {
            # No maintenance, just start explorer.exe
 # Log form data
    Log-FormData -ClientName $clientName -TicketNumber $ticketNumber -Description $description -Pin $pin -MaintenanceDuration $maintenanceDuration

            Run-Explorer
            Close-Windows
        }
    } else {
# Log form data
    Log-FormData -ClientName $clientName -TicketNumber $ticketNumber -Description $description -Pin "BADPIN" -MaintenanceDuration $maintenanceDuration

        [System.Windows.MessageBox]::Show("The PIN is incorrect. Please check your email for the correct PIN.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    }


})

$clear.Add_Click({
    $descriptionTextBox.Clear()
    $clientNameTextBox.Clear()
    $ticketNumberTextBox.Clear()
    $pinTextBox.Clear()
    $maintenanceDurationTextBox.Clear()
})

$cancel.Add_Click({
 [System.Windows.MessageBox]::Show("You will be logged off now.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
 logoff
   Close-Windows
})

# Show the main window
$window.ShowDialog()
