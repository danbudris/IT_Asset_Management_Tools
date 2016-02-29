<#
 
 NOTES:
 12/3/2015
 Need to figure out how to pass information to the AddManager function

 added 'set report path'
 set the initial report path to the $PSSCriptRoot
 added 'export as HTML' option; and figure out how to actually format it decently

 12/8/2015
 added reportpath variable for setting and checking the path the report will be saved in

 added the functionality to name the report as you save it

#>

#LOAD .NET ASSEMBLIES FOR WINFORMS; without this, script won't load outside of ISE
[void]  [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void]  [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

##DEFINE FUNCTIONS

##Sets the location of the self-reported computer configuration information; this should be a full path name using the FQDN
function SetDataPath {

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$tempdata = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the location where your computer inventory information is stored","Set Data Path") 
if ($tempdata -ne $null)
{
$datapath = $tempdata
}
$datapath | Export-Clixml -Path $PSScriptRoot\pathinfo.clixml

}

##Sets the domain controller which will be quiered when any active directory queries are made through the GUI
function SetDomainController{
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$domaincontroller = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the domain controller which will provide AD information","Set DC") 
$domaincontroller | Export-Clixml -Path $PSScriptRoot\dcinfo.clixml }

##Sets the location for reports to be saved
function SetReportPath {

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$reportpath = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the location where computer inventory reports should be created") 
$reportpath | Export-Clixml -Path $PSScriptRoot\reportpath.clixml

}

##Gets the names of the computer report files that were returned by the logonscript distributed with the GPO
function Get-Compnames {
$filenames                = (Get-ChildItem $datapath\*.csv | Select-Object -Property Name -ExpandProperty Name)
foreach ($file in $filenames){

[void] $objListBox.Items.Add("$file")

}#end foreach
}

##Shoes tool strip menu
function OnClick_openToolStripMenuItem($Sender,$AboutMessage){
        [void][System.Windows.Forms.MessageBox]::Show("Version 1.10, Developed by Dan Budris, 2015")
    }

##Shows datapath
function OnClick_ShowDataPath{
$reportpath = Import-Clixml -Path $PSScriptRoot\reportpath.clixml
$datapath = Import-Clixml -Path $PSScriptRoot\pathinfo.clixml
[void][System.Windows.Forms.MessageBox]::Show("Datapath:`n $datapath `n Reportpath: `n $reportpath")
}

#Shows active DC
function OnClick_ShowDC{
$domaincontroller = Import-Clixml -Path $PSScriptRoot\dcinfo.clixml
[void][System.Windows.Forms.MessageBox]::Show("$domaincontroller")
}

function OnClick_ReportOptions{

function ConcatonateCSV
 {

Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
 $SaveFileDialog.initialDirectory = $initialDirectory
 $SaveFileDialog.filter = "All files (*.*)| *.*"
 $SaveFileDialog.ShowDialog() | Out-Null
 $SaveFileDialog.filename
} #end function Get-FileName



$reportname = Get-FileName -initialDirectory $reportpath

$filenames = (Get-ChildItem -path ($datapath+"*.csv") | Select-Object -Property Name -ExpandProperty Name)

foreach ($file in $filenames){

Import-Csv -Path ($datapath + $file) | Export-Csv -Path $reportname -append -force

}#end foreach
}#end function

ConcatonateCSV

}

function OnClick_PingOptions{
        
        $statusLabel.Text = "Pinging...this might take a moment"
        $selectedName = $objname.Text
        $pingresult = ping $selectedName
        [void][System.Windows.Forms.MessageBox]::Show("$pingresult")
        $statusLabel.Text = "Ready"
    }

function OnClick_ADOptions{
    $statusLabel.Text = "Fetching AD info from $domaincontroller...this can take a minute"
    $compname = $objname.text
    $Adresults = Invoke-Command -ComputerName $domaincontroller -ScriptBlock {
    $InsideCompname = $USING:compname
    get-adcomputer -properties * -Filter {name -like $InsideCompname} | select-object -Property Name,DistinguishedName,DNSHostName,Created,IPv4Address,LastLogonDate,ManagedBy,MemberOf,OperatingSystem,PrimaryGroup
    } -ArgumentList $compname | out-string
    
    [void][System.Windows.Forms.MessageBox]::Show("$ADResults")
    $statusLabel.Text = "Ready"
    }


#INSTANTIATE FORMS
#MAIN FORM, LIST BOX
$objForm        = New-Object System.Windows.Forms.Form 
$objLabel       = New-Object System.Windows.Forms.Label
$objListBox     = New-Object System.Windows.Forms.ListBox 

##TEXT BOXES
$objName        = New-Object System.Windows.Forms.RichTextBox
$objMobo        = New-Object System.Windows.Forms.RichTextBox
$objProc        = New-Object System.Windows.Forms.RichTextBox
$objDrive       = New-Object System.Windows.Forms.RichTextBox
$objAntivirus   = New-Object System.Windows.Forms.RichTextBox
$objNetwork     = New-Object System.Windows.Forms.RichTextBox
$objBios        = New-Object System.Windows.Forms.RichTextBox

#BUTTONS
$ResultButton   = New-Object System.Windows.Forms.Button

#STATUS STRIP
$statusStrip    = New-Object System.Windows.Forms.StatusStrip
$statusLabel    = New-Object System.Windows.Forms.ToolStripStatusLabel

#MENUSTRIP, DROPDOWNS
$objMenu        = New-Object System.Windows.Forms.MenuStrip
$objMenuOption  = New-Object System.Windows.Forms.ToolStripMenuItem
$objMenuAbout   = New-Object System.Windows.Forms.ToolStripMenuItem
$PingOption     = New-Object System.Windows.Forms.ToolStripMenuItem
$ADinfo         = New-Object System.Windows.Forms.ToolStripMenuItem
$Refresh        = New-Object System.Windows.Forms.ToolStripMenuItem
$Report         = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutOption    = New-Object System.Windows.Forms.ToolStripMenuItem
$DatapathAbout  = New-Object System.Windows.Forms.ToolStripMenuItem
$datamenu       = New-Object System.Windows.Forms.ToolStripMenuItem
$DCAbout        = New-Object System.Windows.Forms.ToolStripMenuItem
$UpdateDCMenu   = New-Object System.Windows.Forms.ToolStripMenuItem
$CustomReport   = New-Object System.Windows.Forms.ToolStripMenuItem
$UpdateReportPath = New-Object System.Windows.Forms.ToolStripMenuItem

##LOCATION SPACERS
$locationRef    = 150
$locationRef1   = ($locationRef+30)
$locationRef2   = ($locationRef+60)
$locationRef3   = ($locationRef+90)
$locationRef4   = ($locationRef+120)
$locationRef5   = ($locationRef+200)
$locationRef6   = ($locationRef+230)
$locationRef7   = ($locationRef+280)
$locationRef8   = ($locationRef+310)
$locationRef9   = ($locationRef+320)

##MAIN FORM
$objForm.Height           = 550
$objForm.Width            = 500
$objform.MainMenuStrip    = $objMenu
$objform.StartPosition    = "CenterScreen"
$objform.text             = "Computer Hardware Information"

###MAIN FORM LABEL
$objLabel.Location        = New-Object System.Drawing.Size(10,40) 
$objLabel.Size            = New-Object System.Drawing.Size(280,20) 
$objLabel.Text            = "Get Computer Information"

##MAIN LIST BOX
$objListBox.Location      = New-Object System.Drawing.Size(10,60) 
$objListBox.Size          = New-Object System.Drawing.Size(460,20) 
$objListBox.Height        = 120


##STATUS STRIP
$statusLabel.AutoSize     = $true
$statusLabel.text         = "Ready"

##TEXT BOXES FOR COMPUTER INFO RESULTS
##COMPUTER NAME
$objName.Location         = New-Object System.Drawing.Size(10,$locationRef1)
$objName.Size             = New-Object System.Drawing.Size(460,20)
$objName.Height           = 25
$objName.Text             = "NAME"

#COMPUTER MOTHERBOARD
$objMobo.Location         = New-Object System.Drawing.Size(10,$locationRef2)
$objMobo.Size             = New-Object System.Drawing.Size(460,20)
$objMobo.Height           = 25
$objMobo.Text             = "MOTHERBOARD"

##COMPUTER PROCESSOR
$objProc.Location         = New-Object System.Drawing.Size(10,$locationRef3)
$objProc.Size             = New-Object System.Drawing.Size(460,20)
$objProc.Height           = 25
$objProc.Text             = "PROCESSOR"

#COMPUTERS DRIVES
$objDrive.Location        = New-Object System.Drawing.Size(10,$locationRef4)
$objDrive.Size            = New-Object System.Drawing.Size(460,40)
$objDrive.Height          = 60
$objDrive.Text            = "DRIVES"

##COMPUTER ANTIVIRUS
$objAntivirus.Location    = New-Object System.Drawing.Size(10,$locationRef5)
$objAntivirus.Size        = New-Object System.Drawing.Size(460,20)
$objAntivirus.Height      = 25
$objAntivirus.Text        = "ANTIVIRUS"

##COMPUTER NETWORK ADAPTER
$objNetwork.Location      = New-Object System.Drawing.Size(10,$locationRef6)
$objNetwork.Size          = New-Object System.Drawing.Size(460,40)
$objNetwork.Height        = 25
$objNetwork.Text          = "NETWORK CARDS"

##COMPUTER BIOS
$objBios.Location         = New-Object System.Drawing.Size(10,$locationRef7)
$objBios.Size             = New-Object System.Drawing.Size(460,20)
$objBios.Height           = 25
$objBios.Text             = "BIOS"

##RESULT BUTTONS
$ResultButton.Location    = New-Object System.Drawing.Size(225,$locationRef8)
$ResultButton.Size        = New-Object System.Drawing.Size(75,23)
$ResultButton.Text        = "Result"
$ResultButton.Add_Click(
{
$value                    = $objListBox.SelectedItem 

$valueexpanded            = Import-Csv -path $datapath"\"$value   

$objName.Text             = $valueexpanded.Name
$objMobo.Text             = $valueexpanded.Motherboard
$objDrive.Text            = $valueexpanded.Drives
$objProc.Text             = $valueexpanded.Processor
$objAntivirus.Text        = $valueexpanded.Antivirus
$objNetwork.Text          = $valueexpanded.NetworkAdapter
$objBios.Text             = $valueexpanded.BIOS

}
)

##SET DATAPATH FOR COMPUTER INFORMATION

if ((Test-Path $PSScriptRoot\pathinfo.clixml) -ne "true")
{
    SetDataPath
}

$datapath = Import-Clixml -Path $PSScriptRoot\pathinfo.clixml

$domaincontroller = Import-Clixml -Path $PSScriptRoot\dcinfo.clixml

if ((Test-Path $PSScriptRoot\dcinfo.clixml) -ne "true")
{
SetDomainController
}

$domaincontroller = Import-Clixml -Path $PSScriptRoot\dcinfo.clixml


if ((Test-Path $PSScriptRoot\reportpath.clixml) -ne "true")
{
SetReportPath
}

$reportpath = Import-Clixml -Path $PSScriptRoot\reportpath.clixml

##GET INFORMATION FOR LIST BOX

Get-Compnames

##MENUS AND SUBMENUS
##
##MENU STRIP
$objMenu.Location         = New-Object System.Drawing.Point(0,0)
$objMenu.Name             = "Menu"
$objMenu.AutoSize         = "true"
$objMenu.TabIndex         = 0
$objMenu.Text             = "Menu Options"

##INFORMATION MENU
$objMenuAbout.Name        = "Drop Down About"
$objMenuAbout.Size        = New-Object System.Drawing.Size(35,20)
$objMenuAbout.Text        = "&About"

##ABOUT MENU
$AboutOption.Name         = "AboutOption"
$AboutOption.Size         = new-object System.Drawing.Size(152, 22)
$AboutOption.Text         = "&About Program"
$AboutOption.ShortcutKeys = "Control, I"
$AboutOption.ShowShortcutKeys = "true"
$AboutOption.Add_Click( { OnClick_openToolStripMenuItem $AboutOption $EventArgs} )

##DATAPATH ABOUT SUBMENU
$DatapathAbout.Name         = "AboutDatapath"
$DatapathAbout.Size         = new-object System.Drawing.Size(152, 22)
$DatapathAbout.Text         = "&Show Data/Report Paths"
$DatapathAbout.ShortcutKeys = "Control, S"
$DatapathAbout.ShowShortcutKeys = "true"
$DatapathAbout.Add_Click({ OnClick_ShowDataPath })

##DC ABOUT SUBMENU
$DCAbout.Name         = "AboutDC"
$DCAbout.Size         = new-object System.Drawing.Size(152, 22)
$DCAbout.Text         = "&Show Active DC"
$DCAbout.ShortcutKeys = "Control, D"
$DCAbout.ShowShortcutKeys = "true"
$DCAbout.Add_Click({ OnClick_ShowDC })

##OPTIONS MENU
$objMenuOption.Name         = "Drop Down Options"
$objMenuOption.Size         = New-Object System.Drawing.Size(35,20)
$objMenuOption.Text         = "&Options"

##REPORT SUBMENU
$Report.Name                = "ReportOption"
$Report.Size                = new-object System.Drawing.Size(152, 22)
$Report.Text                = "&Produce Computer Report"
$Report.ShortcutKeys        = "Control, Shift, R"
$Report.ShowShortcutKeys    = "true"
$Report.Add_Click({ OnClick_ReportOptions })

##PING SUBMENU
$PingOption.Name          = "PingOption"
$PingOption.Size          = new-object System.Drawing.Size(152, 22)
$PingOption.Text          = "&Ping Computer"
$pingOption.ShortcutKeys  = "Control, P"
$pingOption.ShowShortcutKeys = "true"
$PingOption.Add_Click( {OnClick_PingOptions} )

##AD INFO SUBMENU
$ADinfo.name              = "ADINFO"
$ADinfo.size              = new-object System.Drawing.Size(152,22)
$ADinfo.Text              = "&AD Info"
$ADinfo.ShortcutKeys      = "Control, A"
$ADinfo.ShowShortcutKeys  = "true"
$ADInfo.Add_click({OnClick_ADOptions})

##UPDATE DATA PATH SUBMENU
$datamenu.name            = "DataMenu"
$datamenu.size            = New-Object System.Drawing.Size(152,22)
$datamenu.Text            = "&Update datapath"
$datamenu.ShortcutKeys    = "Control, U"
$datamenu.ShowShortcutKeys= "true"
$datamenu.add_click({SetDataPath})

##UPDATE DC SUBMENU
$UpdateDCMenu.name            = "UpdateDCMenu"
$UpdateDCMenu.size            = New-Object System.Drawing.Size(152,22)
$UpdateDCMenu.Text            = "&Update Active DC"
$UpdateDCMenu.ShortcutKeys    = "Control, D, C"
$UpdateDCMenu.ShowShortcutKeys= "true"
$UpdateDCMenu.add_click({SetDomainController})

##UPDATE REPORTPATH SUBMENU
$UpdateReportPath.name        = "UpdateReportPath"
$UpdateReportPath.size        = New-Object System.Drawing.Size(152,22)
$UpdateReportPath.Text        = "&Update Report Path"
$UpdateReportPath.ShortcutKeys= "Control, U, P"
$UpdateReportPath.ShowShortcutKeys = "true"
$UpdateReportPath.add_click({SetReportPath})

##REFRESH LIST SUBMENU
$Refresh.Name             = "RefreshMenuOption"
$Refresh.size             = New-Object System.Drawing.Size(152,22)
$refresh.Text             = "&Refresh List"
$Refresh.ShortcutKeys     = "Control, R"
$Refresh.ShowShortcutKeys = "true"
$Refresh.add_click({

$datapath = Import-Clixml -Path $PSScriptRoot\pathinfo.clixml

$objListBox.Items.Clear()

Get-Compnames

}) 

##CUSTOM REPORT SUBMENU
$CustomReport.name = "CustomReportMenu"
$CustomReport.size = New-Object System.Drawing.Size(152,22)
$CustomReport.Text = "&Custom AD Stats report"
$CustomReport.ShortcutKeys = "Control, C, R"
$CustomReport.ShowShortCutKeys = "true"
$CustomReport.add_click({CustomReport})

#ADD FORMS TO MAIN FORM
#ADD TO MAIN FORM
$objForm.Controls.Add($objName)
$objForm.Controls.Add($objMobo)
$objForm.Controls.Add($objProc)
$objForm.Controls.Add($objAntivirus)
$objForm.Controls.Add($objNetwork)
$objForm.Controls.Add($objDrive)
$objForm.Controls.Add($objBios)

##MENUS AND DROPDOWNS
$objForm.Controls.Add($objMenu)
$objMenu.Items.AddRange(@($objMenuOption, $objMenuAbout))
$objMenuOption.DropDownItems.AddRange(@($PingOption, $ADinfo, $Refresh, $Report, $datamenu, $updateDCMenu, $UpdateReportPath, $CustomReport))
$objMenuAbout.DropDownItems.AddRange(@($AboutOption, $DatapathAbout, $DCAbout))

##STATUS STRIP
$objform.controls.add($statusStrip)
$statusStrip.Items.Add($statusLabel)

##LABELS AND LISTS
$objForm.Controls.Add($objLabel) 
$objForm.Controls.Add($objListBox)

##BUTTONS
$objForm.Controls.Add($ResultButton)

##SHOW THE FORM
$objForm.ShowDialog()
$objForm.BringToFront()
$objForm.Activate()
$objform.Dispose()





function CustomReport{

 Function ftnChecked 
{

     $CheckedItems = @()

     foreach ($category in $CheckedListBox.CheckedItems) 
     {

         switch ($category.ToString())
         {   
        "Manager"{$CheckedItems += "ManagedBy"}
        "Group Membership"{$CheckedItems += "MemberOf"}
        "Distinguished Name"{$CheckedItems += "DistinguishedName"}
        "DNS Host Name"{$CheckedItems += "DNSHostName"}
        "IPV4 Address"{$CheckedItems += "IPV4Address" }
        "Is Enabled?"{$CheckedItems += "Enabled"}
        "Last Logondate"{$CheckedItems += "LastLogonDate"}
        "Modified"{$CheckedItems += "whenChanged"}
        "Created"{$CheckedItems += "WhenCreated"}
        "Operating System"{$CheckedItems += "OperatingSystem"}
        "Primary Group"{$CheckedItems += "PrimaryGroup"}
        "SID"{$CheckedItems += "SID"}
        "GUID"{$CheckedItems += "objectGUID"}
        }
    }

   
if ($CheckedItems){
         
    $statusLabel.Text = "Fetching AD info from $domaincontroller...this can take a minute"
    $compname = $objname.text
    $adinfostring = Invoke-Command -ComputerName $domaincontroller -ScriptBlock {
    $InsideCompname = $USING:compname
    get-adcomputer -properties $args[1] -Filter {name -like $InsideCompname} | select-object -ExpandProperty $args[1]
    } -ArgumentList $compname,$CheckedItems | out-string
    
    [void][System.Windows.Forms.MessageBox]::Show("$adinfostring")
    $statusLabel.Text = "Ready"
        
    }
        else{
        $adinfo = get-adcomputer -Filter {name -like $compname}
        $adinfostring = $adinfo | out-string
        [void][System.Windows.Forms.MessageBox]::Show($testing)
        }
     }
   
    #CreateForm
    $ReportForm = New-Object -TypeName System.Windows.Forms.Form;
    $ReportForm.Width = 345;
    $ReportForm.Height = 389;
    $ReportForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog;
    $ReportForm.StartPosition = "CenterScreen";
    $ReportForm.MaximizeBox = $false;
    $ReportForm.Text = "Feature selection";

    #CREATE REPORT BUTTON
    $ReportButton = New-Object -TypeName System.Windows.Forms.Button;
    $ReportButton.Text = "Report";
    $ReportButton.Top = $CheckedListBox.Top + $CheckedListBox.Height + 2;
    $ReportButton.Left = ($FeatureForm.Width / 2) - ($OKButton.Width / 2);
    $ReportButton.add_Click({ftnChecked});
    #Add button
    $ReportForm.Controls.Add($ReportButton);

     # Create a CheckedListBox
    $CheckedListBox = New-Object -TypeName System.Windows.Forms.CheckedListBox;
    $CheckedListBox.Width = 325;
    $CheckedListBox.Height = 325;
    $CheckedListBox.Left = 5;
    $CheckedListBox.Top = 5;
    $CheckedListBox.CheckOnClick = $true
    $CheckedListBox.Add_ItemCheck({})
    $ReportForm.Controls.Add($CheckedListBox);

           #Add checkbox items
    $CheckedListBox.Items.Add("Manager") | Out-Null;
    $CheckedListBox.Items.Add("Group Membership") | Out-Null;
    $CheckedListBox.Items.Add("Distinguished Name") | Out-Null;
    $CheckedListBox.Items.Add("DNS Host Name") | Out-Null;
    $CheckedListBox.Items.Add("IPV4 Address") | Out-Null;
    $CheckedListBox.Items.Add("Is Enabled?") | Out-Null;
    $CheckedListBox.Items.Add("Last Logondate") | Out-Null;
    $CheckedListBox.Items.Add("Modified") | Out-Null;
    $CheckedListBox.Items.Add("Created") | Out-Null;
    $CheckedListBox.Items.Add("Operating System") | Out-Null;
    $CheckedListBox.Items.Add("Primary Group") | Out-Null;
    $CheckedListBox.Items.Add("SID") | Out-Null;
    $CheckedListBox.Items.Add("GUID") | Out-Null

    # Show the form
    $ReportForm.ShowDialog();

    

}