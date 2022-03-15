#################################################
# Script created by  : Badrane DERBAZI          #
# Email              : badrane@derbazi.net      #
#                                               #
# For                : Abderrahmane DERBAZI     #
#                                               #
# Date & Time        : 17 SEP 2016 20:02        #
#                                               #
#                                               #
# Description: Exporting Mailbox to a .csv file #
# Script NAme: mmbISE.ps1                       #
# Version    : 1.3                              #
# Licence    : Free                             #
#################################################

###############################################################################################
# Hamzah do not forget to change the  $FolderPath  and the  $File  values:   lines-> 22 & 23  #
#                                     -----------           -----                             #
###############################################################################################

# Folder and File
$FolderPath = "C:\Users\Administrator\OneDrive\Documents\powershell\extractDataFromOutlook"
$File = "MyEmailBox.csv"
$TheFile = Join-Path $FolderPath $File

<#
HO TO EXECUTE THE SCRIPT
************************

OPTION 1
********
1- RUN POWERSHELL ISE
2- OPEN THE FILE MyMailbox.ps1
3- ClICK F5


OPTION 2
********
1- RUN POWERSHELL
2- GO TO THE FOLDER WHERE THE mmb.ps script is located
3- .\MyMailbox.ps1
#>


<#
This script first executed under
--------------------------------
- Windows 8.1 Enterprise Edition
- 16GB of RAM
- PowerShell 4, ISE

Time of Execution : 35s only to read 520 messages
-----------------

AFTER CLOSING THE FORM
the time of killing Outlook Process + Writing to the .csv file is: 1m5s
------------------------------------------------------------------
#>

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$Form = New-Object system.Windows.Forms.Form
$Form.Text = "My Outlook Mailbox v1.3"
$Form.width = 330
$Form.height = 300
$Icon = New-Object system.drawing.icon("$FolderPath\outlook.ico")
$Form.Icon = $Icon 
$form.StartPosition = "CenterScreen"


# CheckBox Sender
$cbx1 = New-Object System.Windows.Forms.CheckBox
$cbx1.Text = "Sender"
$cbx1.AutoSize = $True
$cbx1.Location = New-Object Drawing.Point 10,10
$Form.Controls.Add($cbx1)


$textfield1 = New-Object System.Windows.Forms.TextBox
$textfield1.Location = New-Object Drawing.Point 80,10
$textfield1.Size = New-Object Drawing.Point 180,30
$Form.Controls.Add($textfield1)

# CheckBox Replied
# This check Box can be combined with the others
$cbxRe = New-Object System.Windows.Forms.CheckBox
$cbxRe.Text = "Re"
$cbxRe.AutoSize = $True
$cbxRe.Location = New-Object Drawing.Point 270,12
$Form.Controls.Add($cbxRe)

# CheckBox Received Time
$cbx2 = New-Object System.Windows.Forms.CheckBox
$cbx2.Text = "Received Time"
$cbx2.AutoSize = $True
$cbx2.Location = New-Object Drawing.Point 10,40
$Form.Controls.Add($cbx2)

# Function Pick a Date 1
function pickDate1{
$form = New-Object Windows.Forms.Form 
$form.Text = "Select a Date" 
$form.Size = New-Object Drawing.Size @(243,230) 
$form.StartPosition = "CenterScreen"
$Icon = New-Object system.drawing.icon("$FolderPath\outlook.ico")
$Form.Icon = $Icon 
$calendar = New-Object System.Windows.Forms.MonthCalendar 
$calendar.ShowTodayCircle = $False
$calendar.MaxSelectionCount = 1
$form.Controls.Add($calendar) 
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(38,165)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(113,165)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)
$form.Topmost = $True
$result = $form.ShowDialog() 
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
   {
    $date = $calendar.SelectionStart
    $textfield2.Text = $date
   }
}

# Function Pick a Date 2
function pickDate2{
$form = New-Object Windows.Forms.Form 
$form.Text = "Select a Date" 
$form.Size = New-Object Drawing.Size @(243,230) 
$form.StartPosition = "CenterScreen"
$Icon = New-Object system.drawing.icon("$FolderPath\outlook.ico")
$Form.Icon = $Icon 
$calendar = New-Object System.Windows.Forms.MonthCalendar 
$calendar.ShowTodayCircle = $False
$calendar.MaxSelectionCount = 1
$form.Controls.Add($calendar) 
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(38,165)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(113,165)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)
$form.Topmost = $True
$result = $form.ShowDialog() 
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
   {
    $date = $calendar.SelectionStart
    $textfield4.Text = $date
   }
}

# Date 1
$button1 = New-Object Windows.Forms.Button
$button1.ForeColor = "Blue"
$button1.text = "Date 1"
$button1.Location = New-Object Drawing.Point 125, 35
$button1.Size  = New-Object Drawing.Point 45,25
$button1.add_click({pickDate1})
$form.controls.add($button1)
$textfield2 = New-Object System.Windows.Forms.TextBox
$textfield2.Location = New-Object Drawing.Point 175,40
$textfield2.Size = New-Object Drawing.Point 113,40
$Form.Controls.Add($textfield2)

# Date 2
$button4 = New-Object Windows.Forms.Button
$button4.ForeColor = "Blue"
$button4.text = "Date 2"
$button4.Location = New-Object Drawing.Point 125, 65
$button4.Size  = New-Object Drawing.Point 45,25
$button4.add_click({pickDate2})
$form.controls.add($button4)
$textfield4 = New-Object System.Windows.Forms.TextBox
$textfield4.Location = New-Object Drawing.Point 175,65
$textfield4.Size = New-Object Drawing.Point 113,40
$Form.Controls.Add($textfield4)


# CheckBox Importance
$cbx3 = New-Object System.Windows.Forms.CheckBox
$cbx3.Text = "Importance"
$cbx3.AutoSize = $True
$cbx3.Location = New-Object Drawing.Point 10,70
$Form.Controls.Add($cbx3)

# List of Importance Type 1, 2 or 3
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Size(90,70) 
$comboBox.Size = New-Object System.Drawing.Size(30,20) 
$comboBox.Height = 29
[void] $comboBox.Items.Add("1")
[void] $comboBox.Items.Add("2")
[void] $comboBox.Items.Add("3")
$Form.Controls.Add($comboBox)

# CheckBox Subject
$cbx4 = New-Object System.Windows.Forms.CheckBox
$cbx4.Text = "Subject"
$cbx4.AutoSize = $True
$cbx4.Location = New-Object Drawing.Point 10,100
$Form.Controls.Add($cbx4)

$textfield3 = New-Object System.Windows.Forms.TextBox
$textfield3.Location = New-Object Drawing.Point 80,100
$textfield3.Size = New-Object Drawing.Point 180,30
$Form.Controls.Add($textfield3)

# CheckBox Everything
$cbx5 = New-Object System.Windows.Forms.CheckBox
$cbx5.Text = "Everything"
$cbx5.AutoSize = $True
$cbx5.Location = New-Object Drawing.Point 10,130
$Form.Controls.Add($cbx5)

# Unread Emails
$cbx6 = New-Object System.Windows.Forms.CheckBox
$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
$cbx6.Font = $Font
$cbx6.Text = "Unread"
$cbx6.ForeColor = "Blue"
$cbx6.AutoSize = $True
$cbx6.Location = New-Object Drawing.Point 250,130
$Form.Controls.Add($cbx6)

### Function Get Outlook Messages
Function GetMyeMails{
 $olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type]
 $outlook = new-object -comobject outlook.application
 $namespace = $outlook.GetNameSpace(“MAPI”)
 $folder = $namespace.getDefaultFolder($olFolders::olFolderInBox)
 $folder.items |
 Select-Object -Property Subject, Importance, ReceivedTime, SenderName, UnRead, Body
  
 # Show the Close Button 
  $form.controls.add($button3)
} 

# Export Button
$button2 = New-Object Windows.Forms.Button
$button2.text = "Export"
$button2.Location = New-Object Drawing.Point 10,160
$button2.add_click({GetMyeMails})
$form.controls.add($button2)


# Label 1
$lbl1 = New-Object System.Windows.Forms.Label
$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
$lbl1.Font = $Font
$lbl1.Text = "Combination:"
$lbl1.AutoSize = $True
$lbl1.ForeColor = "Blue"
$lbl1.Location = New-Object Drawing.Point 10,208
$Form.Controls.Add($lbl1)

# Label2
$lbl2 = New-Object System.Windows.Forms.Label
$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Italic)
$lbl2.Font = $Font
$lbl2.Text = "* Sender && Received Time or Sender && Subject"
$lbl2.AutoSize = $True
$lbl2.Location = New-Object Drawing.Point 10,225
$Form.Controls.Add($lbl2)

# Label3
$lbl3 = New-Object System.Windows.Forms.Label
$Fon3 = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Italic)
$lbl3.Font = $Font
$lbl3.Text = "* Unread && Received Time or Unread Only"
$lbl3.AutoSize = $True
$lbl3.Location = New-Object Drawing.Point 10,240
$Form.Controls.Add($lbl3)

# Close Button
$button3 = New-Object Windows.Forms.Button
$button3.text = "Close"
$button3.ForeColor= "Red"
$button3.Location = New-Object Drawing.Point 230,160
$button3.add_click({$form.Close()})


# Show The Form
#$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog()

# Sender or Sender without Replied Messages
if ($cbx1.Checked -eq $true -and $cbxRe.Checked -eq $true){
GetMyeMails | Where {$_.sendername -match $textfield1.Text} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}
ElseIf ($cbx1.Checked -eq $true -and $cbxRe.Checked -eq $false){
GetMyeMails | Where {$_.sendername -match $textfield1.Text -and $_.subject -notmatch "RE:"} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}

# Received Time or Received Time without Replied Messages
if ($cbx2.Checked -eq $True -and $cbxRe.Checked -eq $true){
GetMyeMails | Where {$_.ReceivedTime -ge $textfield2.Text -and $_.ReceivedTime -le $textfield4.Text} | Export-Csv -Path $TheFile-Encoding ASCII -NoTypeInformation  | Out-Default
}
ElseIf ($cbx2.Checked -eq $true -and $cbxRe.Checked -eq $false){
GetMyeMails | Where {$_.ReceivedTime -ge $textfield2.Text -and $_.ReceivedTime -le $textfield4.Text -and $_.subject -notmatch "RE:"} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}

# Importance or Importance without Replied Messages
if ($cbx3.Checked -eq $True -and $cbxRe.Checked -eq $true){
GetMyeMails | Where {$_.Importance -match $comboBox.Text} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}
ElseIf ($cbx3.Checked -eq $true -and $cbxRe.Checked -eq $false){
GetMyeMails | Where {$_.sendername -match $ComboBox.Text -and $_.subject -notmatch "RE:"} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}


# Subject pr Subject without Replied Messages
if ($cbx4.Checked -eq $True -and $cbxRe.Checked -eq $true){
GetMyeMails | Where {$_.Subject -ge $textfield3.Text} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}
ElseIf ($cbx4.Checked -eq $true -and $cbxRe.Checked -eq $false){
GetMyeMails | Where {$_.Subject -ge $textfield3.Text -and $_.subject -notmatch "RE:"} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}

# All Messages or All Messages without Replied Messages
if ($cbx5.Checked -eq $True -and $cbxRe.Checked -eq $true){
GetMyeMails | Where {$_.Subject -ne "@?-+&*()$£TYyvf!!!!@@@@(*&^"} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}
ElseIf ($cbx5.Checked -eq $true -and $cbxRe.Checked -eq $false){
GetMyeMails | Where {$_.Subject -ne "@?-+&*()$£TYyvf!!!!@@@@(*&^" -and $_.subject -notmatch "RE:"} | Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}

# Sender & Received Time
if ($cbx1.Checked -eq $true -and $cbx2.Checked -eq $True) {
GetMyeMails | Where {$_.sendername -match $textfield1.Text -and $_.ReceivedTime -ge $textfield2.Text -and $_.ReceivedTime -le $textfield4.Text} |  Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}

# Sender & Subject
if ($cbx1.Checked -eq $true -and $cbx4.Checked -eq $True) {
GetMyeMails | Where {$_.sendername -match $textfield1.Text -and $_.subject -ge $textfield3.Text} |  Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}


# Unread Messages
if ($cbx6.Checked -eq $true ) {
GetMyeMails | Where {$_.Unread -eq $true} |  Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}

# Unread Messages with Received Time
if ($cbx6.Checked -eq $true -and $cbx2.Checked -eq $True) {
GetMyeMails | Where {$_.Unread -eq $true -and $_.ReceivedTime -ge $textfield2.Text -and $_.ReceivedTime -le $textfield4.Text} |  Export-Csv -Path $TheFile -Encoding ASCII -NoTypeInformation
}

# Release the Outlook Process
kill -processname Outlook*



# 1
######################################
#       Outlook Folders Codes        #
######################################
#						 			 #
# DefaultAppointmentItem 		= 1  #
# DefaultFDefaultderDeletedItems= 3  #
# DefaultFDefaultderOutbox 		= 4  #
# DefaultFDefaultderSentMail 	= 5  #
# DefaultFDefaultderInbox 		= 6  #
# DefaultFDefaultderCalendar 	= 9  #
# DefaultFDefaultderContacts 	= 10 #
# DefaultFDefaultderJournal 	= 11 #
# DefaultFDefaultderNotes 		= 12 #
# DefaultFDefaultderTasks 		= 13 #
# DefaultFDefaultderDrafts 		= 16 #
#									 #
######################################

# 2
######################################
# 		GetDefaultFolder Method      #
######################################
#									 #
# olFolderCalendar					 #
# olFolderContacts					 #
# olFolderDeletedItems				 #
# olFolderDrafts					 #
# olFolderInbox						 #
# olFolderJournal					 #
# olFolderJunk						 #
# olFolderNotes						 #
# olFolderOutbox					 #
# olFolderSentMail					 #
# olFolderTasks						 #
# olPublicFoldersAllPublicFolders	 #
# olFolderConflicts					 #
# olFolderLocalFailures				 #
# olFolderServerFailures			 #
# olFolderSyncIssues			     #
#								     #
######################################

# 3
######################################
# 		 Outlook Item Objects        #
######################################
# https://msdn.microsoft.com/en-us/library/office/ff866278.aspx