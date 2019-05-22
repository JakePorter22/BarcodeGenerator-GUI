#=============================#
# Auth: Jake Porter           #
# Date Made: 5/20/19          #
# Mod By:                     #
# Mod Date:                   #
#=============================#

#=================================================================================#
#                                                                                 #
#     Used by the 10 tablets on the floor to create printable UPC codes           #
#        Also allows the user to change default printer to print to               #
#                                                                                 #
#=================================================================================#



#Functions that creates the form and the code it executes when certain actions happen. 

function GenerateForm
{

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '185,280'
$Form.text                       = "Barcode Generator"
$Form.TopMost                    = $false

#Button that expands the window to select a printer
$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Select Printer"
$Button1.width                   = 60
$Button1.height                  = 38
$Button1.location                = New-Object System.Drawing.Point(59,235)
$Button1.Font                    = 'Microsoft Sans Serif,10'
$Button1.Add_Click({
$Form.ClientSize = '490,280'

})

#Button that starts the print process
$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Print"
$Button2.width                   = 60
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(59,190)
$Button2.Font                    = 'Microsoft Sans Serif,10'
$Button2.Add_Click({
    
    $quantity = $TextBox1.text
    $partsNumber = $TextBox2.text

#Creates the png files of the barcodes using the variables above and stores it in C:\Temp
Set-Location $ZintPath
.\zint.exe --height=60 --notext -o C:\Temp\partsNumber.png -d "$partsNumber"
.\zint.exe --height=60 --notext -o C:\Temp\quantity.png -d "$quantity"


#XAML form that creates the document on which the png and text are put on. 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="MainWindow" Height="11in" Width="8.5in">
    <Grid Name="MyVisual" HorizontalAlignment="Right" Height="11in" Margin="0,0,0,0" VerticalAlignment="Top" Width="8.5in">
        <Grid.RowDefinitions>
            <RowDefinition Height="1.4in"/>
            <RowDefinition Height="1.25in"/>
            <RowDefinition Height="1.5in"/>
            <RowDefinition Height="1.25in"/>
            <RowDefinition Height="1.5in"/>
            <RowDefinition Height="1.25in"/>
            <RowDefinition Height="1.5in"/>
            <RowDefinition Height="1.25in"/>
        </Grid.RowDefinitions>
        <TextBlock Text = "$partsNumber" HorizontalAlignment="Center" Height="Auto"  Grid.Row="0" VerticalAlignment="Center" Width="Auto" FontSize="90" FontWeight="Bold" /> 
        <Image HorizontalAlignment="Center" Height="Auto"  Grid.Row="1" VerticalAlignment="Center" Width="Auto" Stretch="None" Source="C:\Temp\partsNumber.png"/>
        <TextBlock Text = "$quantity" HorizontalAlignment="Center" Height="Auto"  Grid.Row="2" VerticalAlignment="Center" Width="Auto" FontSize="90" FontWeight="Bold" /> 
        <Image HorizontalAlignment="Center" Height="Auto"  Grid.Row="3" VerticalAlignment="Center" Width="Auto" Stretch="None" Source="C:\Temp\quantity.png"/>
        <TextBlock Text = "$partsNumber" HorizontalAlignment="Center" Height="Auto"  Grid.Row="4" VerticalAlignment="Center" Width="Auto" FontSize="90" FontWeight="Bold" /> 
        <Image HorizontalAlignment="Center" Height="Auto"  Grid.Row="5" VerticalAlignment="Center" Width="Auto" Stretch="None" Source="C:\Temp\partsNumber.png"/>
        <TextBlock Text = "$quantity" HorizontalAlignment="Center" Height="Auto"  Grid.Row="6" VerticalAlignment="Center" Width="Auto" FontSize="90" FontWeight="Bold" /> 
        <Image HorizontalAlignment="Center" Height="Auto"  Grid.Row="7" VerticalAlignment="Center" Width="Auto" Stretch="None" Source="C:\Temp\quantity.png"/>
    </Grid>
</Window>
"@
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

#Prints the above document
$print = New-Object System.Windows.Controls.PrintDialog
$print.PrintVisual($MyVisual,"")



})

#Textbox for quantity
$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.TextAlign              = "Center"
$TextBox1.width                  = 100
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(39,150)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

#textbox for parts number
$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.TextAlign              = "Center"
$TextBox2.width                  = 100
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(39,85)
$TextBox2.Font                   = 'Microsoft Sans Serif,10'

#listview for the printers for when you click button1
$ListView1                       = New-Object system.Windows.Forms.ListView
$listView1.View                   = 'Details'
$ListView1.text                  = "listView"
$ListView1.width                 = 180
$ListView1.height                = 150
$ListView1.location              = New-Object System.Drawing.Point(250,35)
$ListView1.Columns.Add("Connected Printers",175) | Out-Null

#adds a line in the listview for each printer connected to on the pc 
Foreach($printer in $printerList)
{
	$results = New-Object System.Windows.Forms.listviewitem($printer)
    $ListView1.Items.Add($results)
    Write-Host $results
}
#switches the default printer depending on what you click in the listview
$ListView1.Add_click({
     foreach($item in $ListView1.SelectedItems){ 
        #Write-Host $Item.text
        $selectPrinter = $Item.text
        $wshNet.SetDefaultPrinter($selectPrinter)
        $defaultPrinter2 = Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true" | select ShareName
        $defaultPrinter2 = ("$defaultPrinter2").substring(12,4)
        $Label5.text = "$defaultPrinter2"
        $Label5.Refresh()
        $Form.ClientSize = '185,280'
     }
 })



#label for part number
$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Part Number"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(55,55)
$Label1.Font                     = 'Microsoft Sans Serif,10'

#label for quantity
$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Quantity"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(58,125)
$Label2.Font                     = 'Microsoft Sans Serif,10'

#label for current printer
$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Current Printer"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(47,10)
$Label4.Font                     = 'Microsoft Sans Serif,10'

#label for what the default printer is deping on the $defaultprinter1 variable
$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "$defaultPrinter1"
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(68,30)
$Label5.Font                     = 'Microsoft Sans Serif,10,style=Bold'

#warning for selecting an item in the listview will change your default printer
$Label6                          = New-Object system.Windows.Forms.Label
$Label6.text                     = "Selecting a Printer will change your default Printer"
$Label6.AutoSize                 = $true
$Label6.width                    = 25
$Label6.height                   = 10
$Label6.location                 = New-Object System.Drawing.Point(215,200)
$Label6.Font                     = 'Microsoft Sans Serif,10,style=Bold,style=Italic'

#adds all of the above to the form
$Form.controls.AddRange(@($ListView1,$TextBox2,$TextBox1,$Button1,$Button2,$Label1,$Label2,$Label4,$label5,$Label6))

$Form.add_Load($OnLoadForm_StateCorrection)
$Form.ShowDialog()| Out-Null


}

# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

#get variables from the start that is used for generating and executing the form Generate
$defaultPrinter1 = Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true" | select ShareName
$defaultPrinter1 = ("$defaultPrinter1").substring(12,4)
$wshNet = New-Object -ComObject WScript.Network
$ZintPath = "C:\Program Files (x86)\Zint"

#gets a list of the printers and selects only the first two letters and the two numbers ex SM46
$printerList = Get-Printer | Select PortName | Select-String -Pattern "SM" 
$printerList = $printerList-replace ".*=" -replace "}.*"

#Generates the form using the function Generate form
GenerateForm

Pause

