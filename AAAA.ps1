#region ADMINCHECK -- Makes sure that the script is launched with admin privileges

param([switch]$Elevated)
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}


$fileserver = "" #Name your file server here ex. \\superserver\

#First tries Powershell 7, then Powershell
if ((Test-Admin) -eq $false)  {
    if (-not $elevated) {        
        try{Start-Process pwsh.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))}
        catch{Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))}
    }
    exit
}
#endregion

#region XAML -- UI For application
$inputXML = @"
<Window x:Class="wpfGUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:AAAA"
        mc:Ignorable="d"
        MinHeight="700" MinWidth="800"
        Title="Adam's Amazing AD Actions" Height="750" Width="800" Background="Black">
    <Grid x:Name="grdMaster" Background="#FF212123">
        <TabControl x:Name="tabMaster">
            <TabItem x:Name="tabUnlockUser" Header="Unlock User">
                <Grid x:Name="grdUnlockUser" Background="#FFE5E5E5">
                    <ListBox x:Name="lstbLockedUsers" Margin="46,138,46,15"/>
                    <Button x:Name="btnRefreshLocked" Content="Refresh" Margin="0,96,340,0" VerticalAlignment="Top" Height="37" FontSize="24" HorizontalAlignment="Right" Width="120"/>
                    <Label x:Name="lblLockedUsers" Content="Locked users:" HorizontalAlignment="Left" Margin="46,99,0,0" VerticalAlignment="Top" FontSize="22"/>
                    <Button x:Name="btnUnlockUser" Content="Unlock User" HorizontalAlignment="Right" Margin="0,100,48,0" VerticalAlignment="Top" Height="36" Width="165" FontSize="18" FontWeight="Bold"/>
                    <Button x:Name="btnUpdate" Content="Check For Updates" HorizontalAlignment="Right" Margin="0,8,55,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblUpdate" Content="" HorizontalAlignment="Right" Margin="0,30,56,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblRefreshed" Content="ERROR" HorizontalAlignment="Right" Margin="0,70,338,0" VerticalAlignment="Top"/>
                    <CheckBox x:Name="chkAutoRefresh" Content="Auto Refresh (30s)" Margin="0,107,218,0" VerticalAlignment="Top" IsChecked="True" HorizontalAlignment="Right" Width="117"/>
                    <Button x:Name="btnTheme" Content="🌗" HorizontalAlignment="Left" Margin="50,30,0,0" VerticalAlignment="Top"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tabCreate" Header="Create New User">
            <Grid x:Name="grdCreate" Background="#FFE5E5E5">
                <ComboBox x:Name="cboSite" HorizontalAlignment="Left" Margin="8,80,0,0" VerticalAlignment="Top" Width="120" Grid.Column="1" MaxDropDownHeight="550" IsDropDownOpen="True"/>
                <Label x:Name="lblSite" Content="Choose a site" HorizontalAlignment="Left" Margin="8,45,0,0" VerticalAlignment="Top" Grid.Column="1" FontWeight="Bold"/>
                <ListView x:Name="lstvCopy" Margin="8,140,0,17" HorizontalAlignment="Left" Width="400" Grid.Column="2">
                    <ListView.View>
                        <GridView>
                            <GridViewColumn DisplayMemberBinding="{Binding Path=Name}" Header="Name" Width="175"/>
                            <GridViewColumn DisplayMemberBinding="{Binding Path=Title}" Header="Title" Width="225"/>
                        </GridView>
                    </ListView.View>
                </ListView>
                <Label x:Name="lblCopy" Content="Optional: Select a user to copy" HorizontalAlignment="Left" Margin="8,110,0,0" VerticalAlignment="Top" Grid.Column="1" FontWeight="Bold"/>
                <TextBox x:Name="txtFirstName" HorizontalAlignment="Right" Margin="0,25,40,0" VerticalAlignment="Top" Width="120"/>
                <TextBox x:Name="txtLastName" HorizontalAlignment="Right" Margin="0,50,40,0" VerticalAlignment="Top" Width="120"/>
                <TextBox x:Name="txtOffice" HorizontalAlignment="Right" Margin="0,75,40,0" VerticalAlignment="Top" Width="120"/>
                <TextBox x:Name="txtDepartment" HorizontalAlignment="Right" Margin="0,100,40,0" VerticalAlignment="Top" Width="120"/>
                <TextBox x:Name="txtTitle" HorizontalAlignment="Right" Margin="0,125,40,0" VerticalAlignment="Top" Width="120"/>
                <TextBox x:Name="txtPhone" HorizontalAlignment="Right" Margin="0,150,40,0" VerticalAlignment="Top" Width="120" MaxLength="10"/>
                <TextBox x:Name="txtExtension" HorizontalAlignment="Right" Margin="0,175,40,0" VerticalAlignment="Top" Width="120" MaxLength = "4"/>
                <TextBox x:Name="txtManager" HorizontalAlignment="Right" Margin="0,200,40,0" VerticalAlignment="Top" Width="120"/>
                <TextBox x:Name="txtDescription" HorizontalAlignment="Right" Margin="0,225,40,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" Height="60"/>
                <Label x:Name="lblFirstName" Content="First Name*:" HorizontalAlignment="Right" Margin="0,20,170,0" VerticalAlignment="Top"/>
                <Label x:Name="lblLastName" Content="Last Name*:" HorizontalAlignment="Right" Margin="0,45,170,0" VerticalAlignment="Top"/>
                <Label x:Name="lblOffice" Content="Office*:" HorizontalAlignment="Right" Margin="0,70,170,0" VerticalAlignment="Top"/>
                <Label x:Name="lblDepartment" Content="Department*:" HorizontalAlignment="Right" Margin="0,95,170,0" VerticalAlignment="Top"/>
                <Label x:Name="lblJobTitle" Content="Job Title*:" HorizontalAlignment="Right" Margin="0,119,170,0" VerticalAlignment="Top"/>
                <Label x:Name="lblPhoneNumber" Content="Phone Number:" HorizontalAlignment="Right" Margin="0,144,170,0" VerticalAlignment="Top"/>
                <Label x:Name="lblExtension" Content="Extension:" HorizontalAlignment="Right" Margin="0,169,170,0" VerticalAlignment="Top"/>
                <Label x:Name="lblManager" Content="Manager:" HorizontalAlignment="Right" Margin="0,193,170,0" VerticalAlignment="Top"/>
                <Label x:Name="lblDescription" Content="Description:" HorizontalAlignment="Right" Margin="0,240,244,0" VerticalAlignment="Top"/>
                <Button x:Name="btnCreateUser" Content="Create User" IsEnabled="False" HorizontalAlignment="Right" Margin="0,427,40,0" VerticalAlignment="Top" Height="64" Width="211" FontSize="36" FontWeight="Bold"/>
                <TextBox x:Name="txtPassword" HorizontalAlignment="Right" Margin="0,379,40,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="195"/>
                <Label x:Name="lblPassword" Content="Password*:" HorizontalAlignment="Right" Margin="0,374,233,0" VerticalAlignment="Top"/>
                <Button x:Name="btnGeneratePW" Content="Generate Password" HorizontalAlignment="Right" Margin="0,350,89,0" VerticalAlignment="Top"/>
                <TextBox x:Name="txtHomeDrive" HorizontalAlignment="Right" Margin="0,324,40,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="174"/>
                <CheckBox x:Name="chkHomeDrive" Content="" HorizontalAlignment="Right" Margin="0,300,89,0" VerticalAlignment="Top" IsChecked="True" RenderTransformOrigin="0.5,0.5"/>
                <Label x:Name="lblHomeDrive" Content="Assign Homedrive" HorizontalAlignment="Right" Margin="0,294,108,0" VerticalAlignment="Top" Width="107"/>
                <Label x:Name="lblUserFolder" Content="Path to user folders:" HorizontalAlignment="Right" Margin="0,319,218,0" VerticalAlignment="Top"/>
            </Grid>
            </TabItem>
            <TabItem x:Name="tabCopy" Header="Copy Groups">
                <Grid x:Name="grdCopy" Background="#FFE5E5E5">
                    <ListBox x:Name="lstbFrom" Margin="0,136,0,0" HorizontalAlignment="Left" Width="283"/>
                    <ListBox x:Name="lstbTo" Margin="0,136,0,0" HorizontalAlignment="Right" Width="283"/>
                    <Label x:Name="lblFrom" Content="Copy From:" HorizontalAlignment="Left" Margin="0,107,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblTo" Content="Copy To:" HorizontalAlignment="Right" Margin="0,108,229,0" VerticalAlignment="Top"/>
                    <CheckBox x:Name="chkDelete" Content="Delete old groups before copying" Margin="0,300,0,0" HorizontalAlignment="Center" Width="198"/>
                    <CheckBox x:Name="chkSelect" Content="Select ALL groups" HorizontalAlignment="Center" Margin="0,327,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <Button x:Name="btnMoveRightCopy" Content="&gt;&gt;" HorizontalAlignment="Center" Margin="0,357,0,0" VerticalAlignment="Top" Height="77" Width="201" FontSize="48"/>
                    <Button x:Name="btnRemoveCopy" Content="Remove" HorizontalAlignment="Center" Margin="0,444,0,0" VerticalAlignment="Top" Height="77" Width="201" FontSize="48"/>
                    <Button x:Name="btnCommitCopy" Content="Commit&#10;Changes" HorizontalAlignment="Center" Margin="0,537,0,0" VerticalAlignment="Top" Height="77" Width="201" FontSize="22" FontWeight="Bold"/>
                    <TextBox x:Name="txtFrom" HorizontalAlignment="Left" Margin="78,113,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200"/>
                    <TextBox x:Name="txtTo" HorizontalAlignment="Right" Margin="0,113,19,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200"/>
                    <Button x:Name="btnClearUsers" Content="Clear Users" HorizontalAlignment="Center" Margin="0,61,0,0" VerticalAlignment="Top" Height="57" Width="157" FontSize="16"/>
                    <CheckBox x:Name="chkCopyDeletedUser" Content="Include Deleted Users (Slower)" HorizontalAlignment="Left" Margin="5,88,0,0" VerticalAlignment="Top"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tabFolders" Header="Folder permissions">
                <Grid x:Name="grdFolders" Background="#FFE5E5E5">
                <Button x:Name="btnSelectFolder" Content="Select Folder" HorizontalAlignment="Left" Margin="15,15,0,0" VerticalAlignment="Top" Height="61" Width="97" FontSize="16" FontWeight="Normal"/>
                <ListBox x:Name="lstbSecurityGroups" Margin="0,154,0,0" HorizontalAlignment="Left" Width="299"/>
                <Label x:Name="lblPermissions" Content="Permissions for" HorizontalAlignment="Left" Margin="0,128,0,0" VerticalAlignment="Top"/>
                <ListBox x:Name="lstbAssignedGroups" Margin="474,327,0,0"/>
                <Button x:Name="btnMoveRight" Content="&gt;" HorizontalAlignment="Left" Margin="327,329,0,0" VerticalAlignment="Top" Height="78" Width="124" FontSize="48" FontWeight="Bold"/>
                <TextBox x:Name="txtSearchUser" HorizontalAlignment="Left" Margin="565,66,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                <ListBox x:Name="lstbSearchAD" Margin="471,91,0,417"/>
                <Label x:Name="lblSearch" Content="Search AD User:" HorizontalAlignment="Left" Margin="469,62,0,0" VerticalAlignment="Top"/>
                <Label x:Name="lblAssignedGroups" Content="Assigned Security Groups:" HorizontalAlignment="Left" Margin="474,302,0,0" VerticalAlignment="Top"/>
                <Button x:Name="btnCommitChanges" Content="Commit&#10;Changes" HorizontalAlignment="Left" Margin="327,425,0,0" VerticalAlignment="Top" Height="78" Width="124" FontSize="18" FontWeight="Bold"/>
                <Button x:Name="btnRemovePermission" Content="Remove" HorizontalAlignment="Left" Margin="327,525,0,0" VerticalAlignment="Top" Height="78" Width="124" FontSize="18" FontWeight="Bold"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tabMigration" Header="User Migration">
                <Grid x:Name="grdMigration" Background="#FFE5E5E5">
                    <TextBox x:Name="txtSourceComp" HorizontalAlignment="Left" Margin="170,39,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <TextBox x:Name="txtDestinationComp" HorizontalAlignment="Left" Margin="170,70,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <TextBox x:Name="txtUser" HorizontalAlignment="Left" Margin="170,98,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" Text="domain\"/>
                    <Label x:Name="lblSourceComp" Content="Source Computer:" HorizontalAlignment="Left" Margin="30,33,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblDestinationComp" Content="Destination Computer:" HorizontalAlignment="Left" Margin="30,64,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblUser" Content="User's account name:" HorizontalAlignment="Left" Margin="30,92,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblCompExample" Content="" HorizontalAlignment="Left" Margin="303,53,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblUserExample" Content="" HorizontalAlignment="Left" Margin="303,93,0,0" VerticalAlignment="Top"/>
                    <CheckBox x:Name="chkTestConnect" Content="Test connection of both computers" HorizontalAlignment="Left" Margin="30,133,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <CheckBox x:Name="chkMusic" Content="End of script tone" HorizontalAlignment="Left" Margin="30,160,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btnMigrate" Content="MIGRATE" HorizontalAlignment="Left" Margin="455,30,0,0" VerticalAlignment="Top" Height="147" Width="290" FontSize="48" FontWeight="Bold" FontFamily="Franklin Gothic Demi"/>
                    <Label x:Name="lblReminder" Content="Starting! Check console for progress." HorizontalAlignment="Left" Margin="124,255,0,0" VerticalAlignment="Top" FontSize="20" FontWeight="Bold" Visibility="Hidden"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tabMapDrive" Header="Remote Drive Map">
                <Grid x:Name="grdMapDrive" Background="#FFE5E5E5">
                    <TextBox x:Name="txtDriveComputerName" Margin="119,32,555,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtDriveLetter" Margin="94,72,0,0" TextWrapping="Wrap" MaxLength="1" VerticalAlignment="Top" HorizontalAlignment="Left" Width="30" IsEnabled="False"/>
                    <TextBox x:Name="txtDrivePath" Margin="169,72,314,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" IsEnabled="False"/>
                    <Label x:Name="lblDriveLetter" Content="Drive Letter:" HorizontalAlignment="Left" Margin="20,67,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblDrivePath" Content="Path:" HorizontalAlignment="Left" Margin="130,67,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblDriveComp" Content="Computer Name:" HorizontalAlignment="Left" Margin="20,26,0,0" VerticalAlignment="Top"/>
                    <ListBox x:Name="lstbUsers" Margin="30,205,0,21" HorizontalAlignment="Left" Width="238"/>
                    <ListBox x:Name="lstbDrives" Margin="326,205,30,21"/>
                    <Label x:Name="lblProfileDrives" Content="User profiles on computer:" HorizontalAlignment="Left" Margin="30,179,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblCurrentDrives" Content="Currently mapped drives:" HorizontalAlignment="Left" Margin="329,179,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btnMapDrive" Content="Map Drive" HorizontalAlignment="Right" Margin="0,80,30,0" VerticalAlignment="Top" Height="54" Width="141" FontSize="22" FontWeight="Bold"/>
                    <Button x:Name="btnSearchComp" Content="Search" HorizontalAlignment="Right" Margin="0,31,509,0" VerticalAlignment="Top"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tabLapsPassword" Header="LAPS Password">
                <Grid x:Name="grdLaps" Background="#FFE5E5E5">
                    <TextBox x:Name="txtLAPSComputerName" Margin="326,77,348,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <Label x:Name="lblLAPSComputerName" Content="Computer Name:" HorizontalAlignment="Left" Margin="228,71,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btnLAPSSearch" Content="Search" HorizontalAlignment="Right" Margin="0,76,302,0" VerticalAlignment="Top"/>
                    <Label x:Name="lblLAPSDisclaimer" Content="This will get the LAPS password of both old LAPS and new" HorizontalAlignment="Left" Margin="225,20,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtLAPSPassword" HorizontalAlignment="Left" Margin="232,117,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="224" IsReadOnly="True"/>
                    <Button x:Name="btnCopyLAPSPW" Content="Copy to Clipboard" HorizontalAlignment="Left" Margin="462,116,0,0" VerticalAlignment="Top"/>
                
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>

"@
#endregion 



#region XAML reading -- This part converts the XAML from above into a display and assigns variables for elements

$global:ReadmeDisplay = $true
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}
catch{
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
  
$xaml.SelectNodes("//*[@Name]") | Foreach-Object{
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
} 

#endregion
$global:hostname = [string](hostname)
$global:whoami = [string](whoami)
$global:whoami = $global:whoami.substring($global:whoami.IndexOf("\")+1,$global:whoami.Length-$global:whoami.IndexOf("\")-1)

Write-Host -ForegroundColor Cyan "Loading Depedencies...."

#region FOCUS CONSOLE -- Disabled, but can be used to hide console (this hides errors too)
#This link shows what each number does
#https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
# Add-Type -Name Window -Namespace Console -MemberDefinition '
# [DllImport("Kernel32.dll")]
# public static extern IntPtr GetConsoleWindow();

# [DllImport("user32.dll")]
# public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
# '
# $global:consolePtr = [Console.Window]::GetConsoleWindow()
function FocusConsole{
    return
}
#     return
#     [Console.Window]::ShowWindow($global:consolePtr, 0)
#     [Console.Window]::ShowWindow($global:consolePtr, 5)
# }
#endregion

#region UNLOCK USER TAB

#Timer for automatic tab refresh at 30 seconds
#Runs in a seperate thread so that the UI still works
$global:timer = New-Object System.Windows.Threading.DispatcherTimer
$global:timer.Interval = [TimeSpan]::FromSeconds(30)
$global:timer.Add_Tick({poplockedusers})
$global:timer.Start()

$WPFchkAutoRefresh.add_Click({
    if ($WPFchkAutoRefresh.IsChecked){
        $global:timer.Start()
    }
    else{
        $global:timer.Stop()
    }
})

#Gets current version of the script by the name, this is a very bad way to do this
#Sets the value to a label to show current version
$wpflblUpdate.Content= $PSCommandPath.Substring($PSCommandPath.LastIndexOf("\")+1,`
                $PSCommandPath.length - 5 - $PSCommandPath.LastIndexOf("\"))

#Sets theme of the program to light or dark, defaults to what your window's app preference is
#This could and probably should be done with a settings file to keep persistence
$global:theme = "Light"
function changeTheme{
    if($global:theme -eq "Light"){
        $maincolor = "#FF212123" #Dark Grey
        $contrastcolor = "White"  
        $contrastcolor2 = "#FF00B503" #Matrix Green
        $accentcolor = "#FF3C3C3C" #Lighter Grey
        $global:theme = "Dark"
    }else{
        $maincolor = "#FFE5E5E5" #Light Grey
        $contrastcolor = "Black"
        $contrastcolor2 = "#FF3C3C3C" #Lighter Grey
        $accentcolor = "White"
        $global:theme = "Light"
    }    
    $labels = Get-Variable "wpflbl*" -ValueOnly
    $labels += Get-Variable "wpfchk*" -ValueOnly
    $grids = Get-Variable "wpfgrd*" -ValueOnly
    $buttons = Get-Variable "wpfbtn*" -ValueOnly
    $textboxes = Get-Variable "wpftxt*" -ValueOnly
    $textboxes += Get-Variable "wpflst*" -ValueOnly
    $tabs = Get-Variable "wpftab*" -ValueOnly
    $others = Get-Variable "wpf*" -ValueOnly -Exclude `
            @("wpflbl*","wpfchk*","wpfgrd*","wpfbtn*","wpftxt*","wpflst*","wpftab*")
    foreach ($label in $labels){
        $label.Foreground = $contrastcolor
    }
    foreach($grid in $grids){
        $grid.background = $maincolor
    }
    foreach ($button in $buttons){
        $button.Background = $maincolor
        $button.Foreground = $contrastcolor
    }
    foreach ($textbox in $textboxes){
        $textbox.background = $accentcolor
        $textbox.foreground = $contrastcolor
    }
    foreach ($tab in $tabs){
        $tab.background = $maincolor   
        $tab.foreground = $contrastcolor2     
    }
    foreach ($other in $others){
        $other.background = $maincolor
        $other.foreground = $contrastcolor2
    }
    $WPFgrdMaster.Background = $contrastcolor2
}
#Toggles theme
$WPFbtnTheme.add_Click({
    changeTheme
})


#Populates the locked user list on Unlock User tab
function poplockedusers(){
    $WPFlstbLockedUsers.items.clear()    
    $lockedusers = Search-ADAccount -LockedOut
    foreach($user in $lockedusers){   
        
        #You can exclude any accounts here, guest should be disabled, but still may appear locked out
        if (-not($user.name -like 'guest')){
            $WPFlstbLockedUsers.items.add($user.name) | Out-Null
        }

    }   

    $time = Get-Date -Format "hh:mm:ss"
    $wpflblRefreshed.content = "Last refreshed: " + $time
    
}


$wpfbtnUpdate.add_Click({
    FocusConsole

    $UpdateVersion = Get-Item "$fileserver\ChangeME!\AAAA*.ps1" #TODO:Change to the script path AAAA*.ps1
    $CurrentVersion = Get-Item $PSCommandPath
    if($UpdateVersion.LastWriteTime -gt $CurrentVersion.LastWriteTime){
        $result = [System.Windows.MessageBox]::Show("Would  you like to update?
        Current Version: $($currentversion.name)
        Update Version: $($updateversion.name)",'AAAA Updater','YesNo')
        if($result -eq 'Yes'){
            #Tries to get the update by running a seperate script,
            #This could probably be done without creating a new script
            $ExecPolic = Get-ExecutionPolicy
            Set-ExecutionPolicy Bypass
            
            #TODO:Change to the script path AAAA*.ps1
            & "$fileserver\ChangeME!\AAAA*.ps1" `
            -Current $PSCommandPath -Update $UpdateVersion.FullName
            Start-Sleep 3
            Set-ExecutionPolicy $ExecPolic
            $form.close()
            Exit
        }
    }else{
        Write-Host "Up to date!" -ForegroundColor Green
    }
})


$WPFbtnUnlockUser.add_Click({
    FocusConsole

    if ($null -eq $WPFlstbLockedUsers.SelectedItem){
        return
    }

    $user = $WPFlstbLockedUsers.SelectedItem
    $user = Get-ADUser -filter 'Name -eq $user'

    try{
        Unlock-ADAccount -identity $user
        Write-Host -ForegroundColor Green Unlocked $user.name !
    }catch{}    

    poplockedusers
    

    $body = @"
    {
        `"time_changed`": `"$(Get-Date -UFormat "%m/%d/%Y %R")`",
        `"type_of_change`":  `"Unlocked User`",
        `"changed_user`":  `"$($user.name)`",
        `"addition_details`":  `"AD Account Unlocked from AAAA`",
        `"ipAddress`":  `"`",
        `"computer`":  `"$global:hostname`",
        `"user`":  `"$global:whoami`"
    }  
"@

    Add-Content -Path $LogsPath -Value "`n$body`n"
})

$WPFbtnRefreshLocked.add_Click({
    poplockedusers
})

#endregion

#region CREATE USER TAB

#TODO:Create Sites and add them to dropdownbox
#Make sure to connect each OU in the switch below
$sites = ("List-Each-OU",`
            "OU2",`
            "OU3"
        )

foreach ($site in $sites){$WPFcboSite.items.add($site) | Out-Null}

$global:OUPath = $null
$global:emailDomain = $null
$global:office = $null

#When site is selected from dropdownbox
$WPFcboSite.add_SelectionChanged({
    $WPFlstvCopy.items.clear()

        #Assigns variables to whatever is selected in dropdown
        #region BIGSWITCH, TODO:List of all sites' OUs as well as email domain
        switch ($WPFcboSite.SelectedItem){
            'List-Each-OU' {$global:OUPath = 'ou=ExampleOU, dc=realplace, dc=local'
                    $global:emailDomain='@realplace.org'
                    $global:office='Real Offices'}
            'OU2'{$global:OUPath = 'ou=OU2, dc=realplace, dc=local'
                    $global:emailDomain='@realplace-butdifferent.org'
                    $global:office='Super Real Offices'}
            'OU3'{$global:OUPath = 'ou=OU3, dc=realplace, dc=local'
                    $global:emailDomain='@realplace.org'
                    $global:office='Not Actually Real Offices'}
            Default{[System.Windows.MessageBox]::Show('Please select a site','Site out of bounds','OK','Error')
                    ; $WPFcboSite.SelectedIndex = 1}
        }
        #endregion

        #Finds users at selected site
        $users = Get-ADUser -Filter * -SearchBase $OUPath -Properties Name, Title, description, manager, memberof, office, department, distinguishedName | Sort-Object Name
        $i=0

        #adds them to array then adds them to listview box
        foreach ($user in $users){
            #puts them into custom object to sort properly in listview
            #this must be done for listview otherwise same value for every column
            $user = [PSCustomObject]@{
                Name = $user.name
                Title = if($user.title -eq $null){$user.description}else{$user.title}
                Manager = $user.manager
                Memberof = $user.memberof
                Office = $user.office
                Department = $user.department
                distinguishedName = $user.distinguishedName
            }            
            $WPFlstvCopy.items.add($user)
            $i++
        }

        $WPFtxtOffice.text = $global:office

        #Header sorting (asc,desc)
        $sb= {
            $col = $_.OriginalSource.Column.Header
            $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlstvCopy.Items)
            $view.SortDescriptions.Clear()
            $sortDescription = New-Object System.ComponentModel.SortDescription($col, 'Ascending')
            $view.SortDescriptions.Add($sortDescription)
        }
    
        $evt = [Windows.RoutedEventHandler]$sb
        $WPFlstvCopy.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $evt)
        
})

#User selected from listview
$WPFlstvCopy.add_SelectionChanged({
    #Clears fields
    $WPFtxtDepartment.text =''
    $WPFtxtTitle.text =''
    $WPFtxtManager.text =''

    #Copies fields from selected user
    $WPFtxtDepartment.text = $WPFlstvCopy.SelectedValue.Department
    $WPFtxtTitle.text = $WPFlstvCopy.SelectedValue.Title
    $WPFtxtOffice.text = $WPFlstvCopy.SelectedValue.Office
    $temp = [string]$WPFlstvCopy.SelectedValue.manager
    $ErrorActionPreference = 'Stop'
    try{$WPFtxtManager.text = $temp.substring(3,$temp.IndexOf(',')-3)}catch{}
})

#Generates random password and adds to pw field based on current season and random symbol
$WPFbtnGeneratePW.add_Click({
    $WPFtxtPassword.text = GeneratePassword
})

 
function GeneratePassword(){
    $wordlist = Get-Content "$fileserver\CHANGEME!\Words.txt" #TODO: Need to have a wordlist for random passwords

    #Symbols to add at the end of pw
    $randsymbols = '!','@','#','$','%','^','&','*',1,2,3,4,5,6,7,8,9,0

    #Weight is how often letters get replaced
    #0.1 is 10% chance, 1 is 100% chance
    $cryptletterweight = 0.1

    #Generates 3 random words to string together
    $words = Get-Random $wordlist -count 3 
    $newwords=@()

    foreach($word in $words){
        $newword = ''
        foreach($letter in $word -split ''){
            $crypt = $null

            #These letters have ($cryptletterweight * 100%)(0.1 * 100% = 10%) chance of being replaced
            #by what $crypt is set to
            switch ($letter){
                'o'{$crypt = '0'}
                'a'{$crypt = '@'}
                'i'{$crypt = '!'}
                't'{$crypt = '7'}
                'e'{$crypt = '3'}
                's'{$crypt = '$'}
                'l'{$crypt = '1'}
            }

            if ($null -ne $crypt){
                #Generates random number, if below threshhold, replaces with crypt (ex. a-->@)
                if ((Get-random -Minimum 0 -Maximum 1.00) -lt $cryptletterweight){
                    $letter = $crypt
                }
            }
            $newword+=$letter
        }
        $newwords += $newword
    }

    $pass =""

    #Adds hypens to password inbetween words
    for ($i =0; $i -lt $newwords.length; $i++){
        if ($i -lt $newwords.length -1){
            $pass += $newwords[$i]+"-"
        }else{
            $pass+= $newwords[$i]
        }
    }

    #Adds 2 characters from $randsymbols to end of password
    $pass += ([string](Get-Random $randsymbols -Count 2)).replace(" ","")

    #recursively finds a password 14 characters or more
    if ($pass.length -lt 14){
        return GeneratePassword
    }else{
        return $pass
    }
}

$WPFtxtFirstName.add_TextChanged({
    validate
})

$WPFtxtLastName.add_TextChanged({
    validate
})

$WPFtxtOffice.add_TextChanged({
    validate
})

$WPFtxtDepartment.add_TextChanged({
    validate
})

$WPFtxtTitle.add_TextChanged({
    validate
    $WPFtxtDescription.text =$WPFtxtTitle.text + " "
})

$WPFtxtDescription.add_TextChanged({
    validate
})

$WPFtxtPassword.add_TextChanged({
    validate
})

function validate{
    #if any of the required fields below are blank, disable button
    if ($WPFtxtFirstName.text -ne ''){
        if ($WPFtxtlastName.text -ne ''){
            if ($WPFtxtOffice.text -ne ''){
                if ($WPFtxtDepartment.text -ne ''){
                    if ($WPFtxtTitle.text -ne ''){
                        if ($WPFtxtPassword.text -ne ''){
                            if ($WPFtxtDescription.text -ne ''){
                                $WPFbtnCreateUser.IsEnabled = $true
                            }else{$WPFbtnCreateUser.IsEnabled = $false}
                        }else{$WPFbtnCreateUser.IsEnabled = $false}
                    }else{$WPFbtnCreateUser.IsEnabled = $false}
                }else{$WPFbtnCreateUser.IsEnabled = $false}
            }else{$WPFbtnCreateUser.IsEnabled = $false}
        }else{$WPFbtnCreateUser.IsEnabled = $false}
    }else{$WPFbtnCreateUser.IsEnabled = $false}
}
$WPFchkHomeDrive.add_Click({
    #Enables/Disables text field
    if($WPFchkHomeDrive.IsChecked){
        $WPFtxtHomeDrive.IsEnabled =$true
    }else{
        $WPFtxtHomeDrive.IsEnabled = $false
    }
})

$WPFbtnCreateUser.add_Click({
    FocusConsole

    #SAM Names deafult to first letter first name + first 4 last name (John Smith = jsmit)
    #Emails default to first letter first name + full last name (John Smith = jsmith@....)

    #Checks pw complexity for length
    if($WPFtxtPassword.text -eq '' -or $WPFtxtPassword.text.length -lt 14){
        [System.Windows.MessageBox]::Show('Please fill out password with correct complexity!','Null fields!','OK','Error')
        return
    }

    #Check SAM name for duplicates
    $samname = ($WPFtxtFirstName.text.Substring(0,1) + $WPFtxtLastName.text.Substring(0,4))
    $samname = $samname.toLower()
    While (Get-ADUser -Filter {samAccountName -eq $samname}){
        Write-Host -ForegroundColor RED "THE CURRENT SAM ACCOUNT NAME $samname IS IN USE"        
        $samname = Read-Host "PLEASE SPECIFY A DIFFERENT SAM NAME"
    }

    #Check UPN for duplicates
    $upn = $WPFtxtFirstName.text.Substring(0,1) + $WPFtxtLastName.text + $global:emailDomain
    While (Get-ADUser -Filter {UserPrincipalName -eq $upn}){
        Write-Host -ForegroundColor RED "THE CURRENT UPN $upn IS IN USE"        
        $upn = Read-Host "PLEASE SPECIFY A DIFFERENT UPN"
    }
    #Will copy OU path of selected user, if there is one
    if ($WPFlstvCopy.selectedindex -gt -1){
        $dn = $WPFlstvCopy.SelectedValue.distinguishedName
        $global:OUPath = $dn.substring($dn.indexof(",")+1,$dn.length - $dn.indexof(",")-1)
    }

    #creates AD User with required fields
    try{
        New-ADUser -givenName $WPFtxtFirstName.text `
            -surname $WPFtxtLastName.text `
            -Name ($WPFtxtFirstName.text + " " + $WPFtxtLastName.text) `
            -DisplayName ($WPFtxtFirstName.text + " " + $WPFtxtLastName.text) `
            -SamAccountName $samname `
            -UserPrincipalName $upn `
            -office $WPFtxtOffice.text `
            -department $WPFtxtDepartment.text `
            -title $WPFtxtTitle.text `
            -email ($WPFtxtFirstName.text.Substring(0,1) + $WPFtxtLastName.text + $global:emailDomain) `
            -description $WPFtxtDescription.text `
            -AccountPassword (ConvertTo-SecureString $WPFtxtPassword.text -AsPlainText -Force) `
            -ChangePasswordAtLogon $false `
            -Enabled $true `
            -ErrorAction Stop `
            -OtherAttributes `
                @{proxyAddresses=("SMTP:" + ($WPFtxtFirstName.Text.Substring(0,1) + $WPFtxtLastName.Text + $global:emailDomain));
                distinguishedName=("CN="+$WPFtxtFirstName.text + " " + $wpftxtlastname.text +","+ $global:OUPath);
                }
    }catch{
        throw
        return
    }
    $newuser = Get-ADUser -Filter 'SamAccountName -eq $samname'

    #Tries to add optional fields to created account
    if ($WPFtxtPhone.text -ne ''){
        Set-ADUser -Identity $newuser.SamAccountName -OfficePhone $WPFtxtPhone.text
    }
    if ($WPFtxtExtension.text -ne ''){
        Set-ADUser -Identity $newuser.SamAccountName -Replace @{'ipPhone'=$WPFtxtExtension.text}
    }
    if ($WPFtxtManager.text -ne ''){
        $manager = ($WPFtxtManager.Text).Trim()
        $manager = Get-ADUser -Filter 'Name -like "$manager*"'
        Set-ADUser -Identity $newuser.SamAccountName -Manager $manager.distinguishedName
    }

    #region create homedrive
    if ($WPFchkHomeDrive.IsChecked){
        #Creates the folder
        $folderpath = $WPFtxtHomeDrive.Text
        New-Item -path $folderpath -ItemType "Directory" -name $samname
        $folderpath = $folderpath + "\"+$samname

        #Set user's homedrive to created folder
        Set-ADUser -Identity $newuser.SamAccountName -HomeDirectory $folderpath -HomeDrive "U:"
        
        #Removes all permissions from created folder
        $acl = Get-ACL -path $folderpath    
        $permissions = $acl.access
        foreach($permission in $permissions){
            $acl.RemoveAccessRule($permission)
        }

        #Adds only these permissions to the folder

        $permissions = @()
        $permissions += [PSCustomObject]@{
            IdentityReference = 'NT AUTHORITY\SYSTEM'
            FileSystemRights = 'FullControl'
            InheritanceFlags = 'ContainerInherit,ObjectInherit'
            PropogationFlags = 'None'
            AccessControlType = 'Allow'
        }
        $permissions += [PSCustomObject]@{
            IdentityReference = 'DOMAIN\Domain Admins' #TODO:Change domain name
            FileSystemRights = 'FullControl'
            InheritanceFlags = 'ContainerInherit,ObjectInherit'
            PropogationFlags = 'None'
            AccessControlType = 'Allow'
        }
        $permissions += [PSCustomObject]@{
            IdentityReference = "DOMAIN\$samname" #TODO:Change domain name
            FileSystemRights = 'Modify'
            InheritanceFlags = 'ContainerInherit,ObjectInherit'
            PropogationFlags = 'None'
            AccessControlType = 'Allow'
        }
        foreach($permission in $permissions){
            $rule = New-Object System.Security.accesscontrol.filesystemaccessrule(`
            $permission.IdentityReference,`
            $permission.FileSystemRights,`
            $permission.InheritanceFlags,`
            $permission.PropogationFlags,`
            $permission.AccessControlType)
            $acl.SetAccessRule($rule)
        }

        #Sets owner of folder to domain admins
        $owner = New-Object -TypeName System.Security.Principal.NTAccount('domain\Domain Admins')#TODO:Change domain name
        $acl.SetOwner($owner)
        $acl.SetAccessRuleProtection($true,$false)#-- DISABLES INHERITANCE
        Set-ACL -path $folderpath -AclObject $acl
    }
    #endregion
    if($WPFlstvCopy.SelectedIndex -gt -1){
        $WPFlstvCopy.SelectedValue.Memberof | Add-ADGroupMember -Members $newuser
        Write-Host -ForegroundColor Green $newuser.name " added! Copied groups from " $WPFlstvCopy.SelectedValue.Name
    }else{Write-Host -ForegroundColor Green $newuser.name " added! No groups copied."}
    Write-Host -ForegroundColor Green "Password is set to: " -NoNewLine
    Write-Host -ForegroundColor White $WPFtxtPassword.text -BackgroundColor Black
    
    $body = @"
    {
        `"time_changed`": `"$(Get-Date -UFormat "%m/%d/%Y %R")`",
        `"type_of_change`":  `"New AD User`",
        `"changed_user`":  `"$($newuser.name)`",
        `"addition_details`":  `"Department: $($WPFtxtDepartment.text); Title: $($WPFtxtTitle.text); Phone: $($WPFtxtPhone.text); Extension: $($WPFtxtExtension.text); Manager: $($WPFtxtManager.text); Copied From: $($WPFlstvCopy.SelectedValue.Name)`",
        `"ipAddress`":  `"`",
        `"computer`":  `"$global:hostname`",
        `"user`":  `"$global:whoami`"
    }  
"@

    Add-Content -Path $LogsPath -Value "`n$body`n"

    #Clears text when user is created, then displays popup
    $WPFtxtFirstName.text = ""
    $WPFtxtLastName.text = ""
    $WPFtxtDepartment.text = ""
    $WPFtxtTitle.text = ""
    $WPFtxtPassword.text = ""
    $WPFtxtPhone.text = ""
    $WPFtxtExtension.text = ""
    $WPFtxtManager.text = ""
    $WPFtxtDescription.text = ""
    $wpflstbusers.SelectedIndex = -1
    [System.Windows.MessageBox]::Show('User created!','Success')

    
})
#endregion

#region COPY GROUPS TAB
$global:gettinguserfrom=$true
$global:gettinguserto=$true

$WPFchkDelete.IsEnabled = $false
$wpftxtFrom.add_TextChanged({
    if ($global:gettinguserfrom){
        $wpflstbFrom.items.clear()
        $temp = $wpftxtFrom.Text
        if ($WPFchkCopyDeletedUser.IsChecked){
            if ($wpftxtfrom.text.length -gt 3){
                $users = Get-ADObject -Filter "Name -like '*$temp*'" -IncludeDeletedObjects -Properties Name,Displayname,Memberof `
                        | Where-Object ObjectClass -eq 'user'
            }
        }else{
            if ($wpftxtfrom.text.length -gt 2){
                $users = Get-ADUser -Filter "Name -like '*$temp*'" -Properties DisplayName, Memberof
            }
        }
        $WPFlstbSearchAD.items.clear()
        $global:fromusergroups =@()
        foreach ($user in $users){
            if (-not $user.displayname -eq ''){
                $global:fromusergroups+= [PSCustomObject]@{
                    Displayname = $user.displayname
                    Groups = $user.memberof
                }
                $wpflstbFrom.items.add($user.DisplayName)
            }
        }
    }
})
$global:fromgroups=@()
$WPFlstbFrom.add_SelectionChanged({
    if ($global:gettinguserfrom){
        $global:gettinguserfrom=$false
        $user = [string]$WPFlstbFrom.SelectedItem

        $user =  Get-ADObject -Filter "Name -like '*$user*'" -IncludeDeletedObjects -Properties Memberof, DisplayName

        $temp2 = $user.memberof
        $global:fromgroups.clear()
        $WPFlstbFrom.items.clear()
        $wpftxtFrom.text=$user.Displayname
        foreach ($group in $temp2){
            $global:fromgroups+=$group
            $temp3 = [string] $group
            $temp3 = $temp3.substring(3,$temp3.IndexOf(',')-3)
            $WPFlstbFrom.items.add($temp3)
        }         
    }
})

$WPFlstbTo.add_SelectionChanged({
    if ($global:gettinguserTo){
        $global:gettinguserTo=$false
        $user = [string]$WPFlstbTo.SelectedItem   
        $user = Get-ADUser -Filter "Name -like '$user'" -Properties DisplayName
        $temp2 = Get-ADPrincipalGroupMembership -Identity $user 
        $Togroups=@()
        $Togroups.clear()
        $WPFlstbTo.items.clear()
        $wpftxtTo.text=$user.Displayname
        foreach ($group in $temp2){
            $Togroups+=$group
            $temp3 = [string] $group
            $temp3 = $temp3.substring(3,$temp3.IndexOf(',')-3)
            $WPFlstbTo.items.add($temp3)
        }         
    }
})

$WPFtxtTo.add_TextChanged({
    if ($global:gettinguserTo){
        $wpflstbTo.items.clear()
        $temp = $wpftxtTo.Text
        if ($wpftxtto.text.length -gt 2){
            $users = Get-ADUser -Filter "Name -like '*$temp*'"
        }
        $WPFlstbSearchAD.items.clear()
        foreach ($user in $users){
            $wpflstbTo.items.add($user.Name)
        }
    }
})
$WPFbtnMoveRightCopy.add_Click({
    if ($WPFchkSelect.IsChecked){
        foreach ($group in $global:fromgroups){
            $group = [string]$group
            $group = $group.substring(3,$group.IndexOf(',')-3)
            if (-not $WPFlstbTo.items.Contains($group)){
                $WPFlstbTo.items.add($group)
                $WPFlstbFrom.items.remove($group)
            }            
        }
    }else{
        $group = $WPFlstbFrom.SelectedItem
        if (-not $WPFlstbTo.items.Contains($group)){
            $WPFlstbTo.items.add($group)
        }
    }

})
$WPFbtnRemoveCopy.add_Click({
    $WPFlstbTo.Items.Remove($WPFlstbTo.SelectedItem)
})
$WPFbtnCommitCopy.add_Click({
    FocusConsole
    $user = [string]$wpftxtTo.text
    $user = Get-ADUser -Filter 'Name -eq $user'
    foreach ($group in $wpflstbTo.items){
        $group = Get-ADGroup -Filter 'Name -like $group'
        try{Add-ADGroupMember -Identity $group -Members $user}
        catch{}        
    }
    Write-Host -ForegroundColor Green Done!
    $body = @"
    {
        `"time_changed`": `"$(Get-Date -UFormat "%m/%d/%Y %R")`",
        `"type_of_change`":  `"Copied Groups`",
        `"changed_user`":  `"$($user.name)`",
        `"addition_details`":  `"Added these groups: $($wpflstbTo.Items)`",
        `"ipAddress`":  `"`",
        `"computer`":  `"$global:hostname`",
        `"user`":  `"$global:whoami`"
    }  
"@

    Add-Content -Path $LogsPath -Value "`n$body`n"
})
$WPFchkSelect.add_Click({
    if ($WPFchkSelect.IsChecked){
        $WPFbtnMoveRightCopy.content='>>'
    }else{
        $WPFbtnMoveRightCopy.content='>'
    }
})
$WPFbtnClearUsers.add_Click({
    $global:gettinguserfrom=$true
    $global:gettinguserto=$true
    $WPFlstbFrom.items.clear()
    $WPFlstbTo.items.clear()
    $WPFtxtTo.Text=''
    $wpftxtFrom.Text=''
})
#endregion

#region FOLDERS TAB
$wpfbtnSelectFolder.add_Click({
    FocusConsole
    $wpflstbSecurityGroups.items.Clear()
    Add-Type -AssemblyName System.Windows.Forms
    $Folder = New-Object System.Windows.Forms.FolderBrowserDialog
    
    if($Folder.ShowDialog() -eq 'Cancel'){
        Write-Host -ForegroundColor Red "Please select a folder!"
        return
    }
    Write-Host 'Folder: ' $Folder.SelectedPath
    $permissions = (Get-Acl $Folder.SelectedPath).Access | Select-Object *
    $displayfolder = [string]$folder.SelectedPath
    $displayfolder = $displayfolder.substring($displayfolder.LastIndexOf("\") + 1,$displayfolder.length-$displayfolder.LastIndexOf("\") -1 )
    $WPFlblPermissions.Content = "Permissions for " + $displayfolder + " :"

    Foreach ($permission in $permissions){
        if ($permission.IdentityReference -like "*domain admin*" -or $permission.IdentityReference -like "*lab*"){} #Skips over these
        else{
            $permission = [string]$permission.IdentityReference
            $permission = $permission.substring($permission.IndexOf('\')+1,$permission.length - $permission.indexof('\')-1)
            $wpflstbSecurityGroups.items.add($permission)}
    }
    
    $wpflstbSecurityGroups.items.add("------SUB FOLDERS' EXPLICIT PERMS----------")

    
    $subfolders = Get-ChildItem -Path $Folder.SelectedPath 
    foreach ($subfolder in $subfolders){
        $permissions = (Get-Acl $subfolder).Access | Select-Object *
        Foreach ($permission in $permissions){
            if ($permission.IdentityReference -like "*domain admin*" -or $permission.IdentityReference -like "*lab*"){} #Skips over these
            else{
                $permission = [string]$permission.IdentityReference
                $permission = $permission.substring($permission.IndexOf('\')+1,$permission.length - $permission.indexof('\')-1)    
                if (-not $wpflstbSecurityGroups.items.Contains($permission)){
                    $output = [string]$subfolder
                    $wpflstbSecurityGroups.items.add("Perms for " + ($output).substring($output.LastIndexOf("\")+1,$output.length-$output.LastIndexOf("\")-1)+':')
                    $wpflstbSecurityGroups.items.add($permission)
                }
            }
        }
    }
})

$WPFtxtSearchUser.add_TextChanged({
    $users = $WPFtxtSearchUser.Text
    $users = Get-ADUser -Filter "Name -like '*$users*'"
    
    $WPFlstbSearchAD.items.clear()
    foreach ($user in $users){
        $WPFlstbSearchAD.items.add($user.Name)
    }
})

$global:newgroups = @()
$global:selecteduser = $null

$WPFlstbSearchAD.add_SelectionChanged({
    $WPFlstbAssignedGroups.items.clear()
    $user = [string]$WPFlstbSearchAD.SelectedItem   
    $global:selecteduser = $user
    $user = Get-ADUser -Filter {Name -like $user}
    $temp2 = Get-ADPrincipalGroupMembership -Identity $user
    $global:newgroups = @()
    foreach ($group in $temp2){
        $global:newgroups+=$group
        $temp3 = [string] $group
        $temp3 = $temp3.substring(3,$temp3.IndexOf(',')-3)
        $WPFlstbAssignedGroups.items.add($temp3)
    }
})


$WPFbtnMoveRight.add_Click({    
    $temp = [string]$WPFlstbSecurityGroups.SelectedItem
    $temp = $temp.substring($temp.IndexOf('\')+1,$temp.Length-$temp.IndexOf('\')-1)
    if (-Not $WPFlstbAssignedGroups.items.Contains($temp)){
        $WPFlstbAssignedGroups.items.add($temp)
    }else{
        Write-Host -ForegroundColor Yellow Already contains $temp
    }
    $group = Get-ADGroup -Identity $temp
    $global:newgroups+=$group
})

$WPFbtnRemovePermission.add_Click({
    $temp = $WPFlstbAssignedGroups.SelectedItem
    $group = Get-ADGroup -Identity $temp
    for($i=0; $i -lt $global:newgroups.count;$i++){
        if ([string]$global:newgroups[$i] -like [string]$group){
            $global:newgroups[$i] = $null
        }
    }
    $WPFlstbAssignedGroups.items.Remove($WPFlstbAssignedGroups.SelectedItem)
    $WPFlstbAssignedGroups.SelectedIndex = $WPFlstbAssignedGroups.TabIndex
})

$WPFbtnCommitChanges.add_Click({   
    FocusConsole
    $user = Get-ADUser -Filter "Name -eq '$global:selecteduser'"
    Write-Host User: $user
    for ($i=0;$i -lt $global:newgroups.count;$i++){
        $group = [string]$global:newgroups[$i]
        write-host ADDING TO GROUP : $group
        $group = Get-ADGroup -Identity $group
        Try{Add-ADGroupMember -Identity $group -Members $user}
        Catch{}
    }
    $global:newgroups = @()
    $body = @"
    {
        `"time_changed`": `"$(Get-Date -UFormat "%m/%d/%Y %R")`",
        `"type_of_change`":  `"Folder permissions added`",
        `"changed_user`":  `"$($user.name)`",
        `"addition_details`":  `"Added group $group from a folder`",
        `"ipAddress`":  `"`",
        `"computer`":  `"$global:hostname`",
        `"user`":  `"$global:whoami`"
    }  
"@
    Add-Content -Path $LogsPath -Value "`n$body`n"
})
#endregion

#region USER MIGRATION -- This section relies on User State Migration Tool (USMT) https://learn.microsoft.com/en-us/windows/deployment/usmt/usmt-overview
$WPFbtnMigrate.add_Click({
    FocusConsole
    Add-Type -AssemblyName System.Speech
    $TTS = New-Object System.Speech.Synthesis.SpeechSynthesizer
    
    $sourceComputer = $WPFtxtSourceComp.Text
    $destinationComputer = $WPFtxtDestinationComp.Text
    
    if ($WPFtxtSourceComp.text -eq '' -or $WPFtxtDestinationComp.text -eq ''){
        Write-Host -ForegroundColor Red "PLEASE FILL IN FIELDS!"
        return
    }else{
        $WPFlblReminder.Visibility = "Visible"
    }

    #region test-connection   -- Tests connection to each computer before trying to migrate user between computers

    if ($WPFchkTestConnect.IsChecked){
        Write-Host Testing connection to $sourceComputer....
        if (-not(Test-Connection -TargetName $sourceComputer -Quiet)){
            Write-Host -ForegroundColor Red Could not ping $sourceComputer, cancelling migration.
            return
        }
        Write-Host -ForegroundColor Green Success!  
        Write-Host Testing connection to $destinationComputer.... 
        if (-not(Test-Connection -TargetName $destinationComputer -Quiet)){
            Write-Host -ForegroundColor Red Could not ping $destinationComputer, cancelling migration.
            return
        }
        Write-Host -ForegroundColor Green Success!    
    }

    #endregion

    $username = $WPFtxtUser.Text
   
    Write-Host $username

    #region scan+load -- scans + saves from one computer, copies to the other, loads + applies from the other
    $before = Get-Date
    Write-Host (Get-Date)"Copying files to $sourceComputer (This file is 43.7MB)"
    Copy-Item "C:\USMT\USMTFiles" -Destination "\\$sourcecomputer\c$\" -Recurse

    Write-Host (Get-Date)"Starting scanstate (This can take a while)"
    Invoke-Command -ComputerName $sourceComputer -ScriptBlock{
        C:\USMTFiles\scanstate.exe "C:\USMTFiles" /i:c:\usmtfiles\MigUser.xml /i:c:\usmtfiles\MigDocs.xml /i:c:\usmtfiles\MigApp.xml /ue:*\* /ui:$Using:username /v:1 /c /o /efs:copyraw
    }

    $filesize = (Get-ChildItem "\\$sourcecomputer\c$\USMTFiles\USMT" | Measure-Object -Property Length -Sum).Sum / 1GB
    Write-Host (Get-Date)"Copying files to $destinationComputer (This file is $filesize GB)"
    Copy-Item "\\$sourcecomputer\c$\USMTFiles" -Destination "\\$destinationComputer\c$\" -Recurse

    Write-Host Starting loadstate
    Invoke-Command -ComputerName $destinationComputer -ScriptBlock{
        C:\USMTFiles\loadstate.exe "C:\USMTFiles" /i:c:\usmtfiles\MigUser.xml /i:c:\usmtfiles\MigDocs.xml /i:c:\usmtfiles\MigApp.xml /v:1 /c
    }
    
    Remove-Item -Path "\\$sourcecomputer\c$\USMTFiles" -Recurse
    Remove-Item -Path "\\$destinationComputer\c$\USMTFiles" -Recurse

    Write-Host -ForegroundColor Green "Done!`nTotal Time:"
    $totaltime = New-Timespan -Start $before -End (Get-Date)
    Write-Host $totalTime

    #This section is a fun high scores for whoever migrated a computer the fastest
    Try{
        $ErrorActionPreference = 'Stop'

        $mytime = $totalTime.TotalSeconds
        
        $highscores= Import-CSV "\\server\share\scores.csv" | Sort-Object {[int]$_.Time} #TODO: Change server to something applicable

        $placed = $false

        if ($mytime -lt $highscores[0].Time){
            write-host -ForegroundColor Magenta "**********************************"
            Write-Host -NoNewLine -ForegroundColor Blue "**********"
            Write-Host -NoNewline -ForegroundColor Green "NEW HIGH SCORE"
            Write-Host -ForegroundColor Cyan "**********"
            Write-Host -ForegroundColor DarkYellow "**********************************"
            Write-Host "Your time was $mytime on" (Get-Date)
            $placed = $true
        }else{
            $i = 1
            foreach ($score in $highscores){
                $seconds = $score.time % (24 * 3600)
                $hours = [math]::Floor($seconds / 3600)
                $seconds %= 3600
                $minutes =  [math]::Floor($seconds / 60)
                $seconds %= 60
                
                if ($i -eq 1){
                    Write-Host $score.User"is currently in 1st place with a time of $hours hours, $minutes minutes, `
                    and $seconds seconds while copying"$score.source"to"$score.destination"`
                    with a file size of"$score.FileSize"on"$score.date"!"
                }
                if ($mytime -lt [int]$score.time){
                    Write-Host "You have placed "$i
                    $placed = $true
                    break
                }
                $i++
            }
        }
        if (!$placed){
            Write-Host -ForegroundColor DarkRed "You're last place at "$i
        }
        $me = whoami.exe
        "$myTime,$me,$sourcecomputer,$destinationcomputer,$(Get-Date),$filesize" `
            | Add-Content "\\server\share\scores.csv" #TODO: Change server to something applicable
            
    }catch{throw}

    #endregion
    $body = @"
    {
        `"time_changed`": `"$(Get-Date -UFormat "%m/%d/%Y %R")`",
        `"type_of_change`":  `"USMT`",
        `"changed_user`":  `"$username`",
        `"addition_details`":  `"User migrated from $sourcecomputer to $destinationcomputer in $mytime with a filesize of $filesize`",
        `"ipAddress`":  `"`",
        `"computer`":  `"$global:hostname`",
        `"user`":  `"$global:whoami`"
    }  
"@
    Add-Content -Path $LogsPath -Value "`n$body`n"

    #This either plays the tetris theme or speaks and says migration complete, comment out if you want neither
    #End of script tetris music
    if ($wpfchkMusic.IsChecked){
        $one = @((1320,500),(1188,250),(1056,250),(990,750),(1056,250),(1188,500),(1320,500),(1056,500),(880,500),(880,500))
        foreach ($N in $one) { 
            [Console]::Beep($N[0],$N[1]) 
        }
    }else{
        $TTS.Speak("MIGRATION COMPLETE")
    }
})
#endregion

#region MAP DRIVES -- Remotely map drive using registry

$global:mappingcomputer = $null
$global:sids = $null

$WPFbtnSearchComp.add_Click({
    FocusConsole
    $wpflstbusers.items.clear()
    $WPFlstbDrives.items.clear()
    $global:mappingcomputer= $WPFtxtDriveComputerName.text

    if (Test-Connection -ComputerName $global:mappingcomputer){
        $WPFtxtDriveLetter.IsEnabled = $true
        $WPFtxtDrivePath.IsEnabled = $true

        #Have to save option to CSV, then transfer to local computer
        #There is probably a way to get the variable to come out of the invoke

        Invoke-Command -ComputerName $global:mappingcomputer -ScriptBlock {
            New-Item -Path "c:\" -Name "drivetemp" -ItemType "directory"
            try{
                $users = Get-ChildItem -Path "REGISTRY::HKEY_USERS"
            }catch{}
            $output=@()
            $FormatEnumerationLimit = -1
            foreach ($user in $users){
                $drives = Get-ChildItem -Path "REGISTRY::$user\Network" -ErrorAction SilentlyContinue
                $drivepath = @()
                foreach ($drive in $drives){
                    $drive = [string]$drive
                    $drive = $drive.substring($drive.length - 1,1)
                    $drivepath += Get-ItemProperty -Path "REGISTRY::$user\network\$drive"
                }
                $temp = [PSCustomObject]@{           
                    Name = $user
                    Drives = [string]$drives
                    DrivePath = [string]$drivepath.remotepath
                }
                $output+=$temp
            }
            $output = $output | Export-CSV -Path "C:\drivetemp\users.csv"
        }

        $global:sids = Import-CSV -path "\\$global:mappingcomputer\c$\drivetemp\users.csv"
        foreach ($sid in $global:sids){
            $name = [string]$sid.name
            $name = $name.substring($name.indexof("\")+1,$name.length-$name.indexof("\")-1)            
            try{$user=Get-ADUser -Filter "SID -eq '$name'" -Properties Name,SID}catch{}
            if ($user -ne $null){
                $wpflstbusers.items.add($user.name)
            }            
        }
        Remove-Item -path "\\$global:mappingcomputer\c$\drivetemp" -Recurse
    }
})

$wpflstbusers.add_SelectionChanged({
    $WPFlstbDrives.items.clear()
    if($wpflstbusers.SelectedIndex -eq -1){
        return
    }
    $user = $wpflstbusers.SelectedItem
    $user = Get-ADUser -filter "Name -eq '$user'"
    $user = $user.sid
    $sid = $global:sids
    for ($i = 0; $i -lt $sid.length;$i++){
        if ($sid[$i].name -like "*$user*"){
            $drives = $sid[$i].drives -split " "
            $paths = $sid[$i].drivepath -split " \",0, "SimpleMatch"
            for($j = 0; $j -lt $drives.length; $j++){
                $drive = [string]$drives[$j]
                $path = $paths[$j]
                try{$drive = $drive.substring($drive.LastIndexOf("\")+1,1)
                    $WPFlstbDrives.items.add($drive+":" + $path)
                }catch{}
                
            } 
        }
    }
})

$WPFtxtDriveComputerName.add_SelectionChanged({
    $WPFtxtDriveLetter.IsEnabled = $false
    $WPFtxtDrivePath.IsEnabled = $false
})

$WPFbtnMapDrive.add_Click({
    FocusConsole
    $driveletter = ($WPFtxtDriveLetter.text).toUpper()+":"
    foreach($item in $WPFlstbDrives.items){
        $item = [string]$item
        $item = $item.substring(0,2)
        if ($driveletter -eq $item -or $driveletter -eq 'C:'){
            Write-Host -ForegroundColor Red "Drive letter already mapped!"
            return
        }elseif($driveletter -eq ":" -or $driveletter -eq ''){
            Write-Host -ForegroundColor Red "Please enter a drive letter!"
            return
        }
    }

    Write-Host "Mapping "$driveletter...
    $sid = $wpflstbusers.selecteditem
    $sid = Get-ADUser -Filter "Name -eq '$sid'"
    $sid = $sid.sid
    $path = $WPFtxtDrivePath.text

    Invoke-Command -computer $global:mappingcomputer -ScriptBlock{
        $sid = $Using:sid
        $driveletter= $using:driveletter
        $driveletter= $driveletter.substring(0,1)
        $DrivePath = $Using:path
        $Path = "REGISTRY::HKEY_USERS\$sid\Network"
        New-Item -Path $path -Name $DriveLetter
        New-ItemProperty -Path $Path\$DriveLetter\ -Name ConnectFlags -PropertyType DWORD -Value 0
        New-ItemProperty -Path $Path\$DriveLetter\ -Name ConnectionType -PropertyType DWORD -Value 1
        New-ItemProperty -Path $Path\$DriveLetter\ -Name DeferFlags -PropertyType DWORD -Value 4
        New-ItemProperty -Path $Path\$DriveLetter\ -Name ProviderName -PropertyType String -Value "Microsoft Windows Network"
        New-ItemProperty -Path $Path\$DriveLetter\ -Name ProviderType -PropertyType DWORD -Value 131072
        New-ItemProperty -Path $Path\$DriveLetter\ -Name RemotePath -PropertyType String -Value "$DrivePath"
        New-ItemProperty -Path $Path\$DriveLetter\ -Name UserName -PropertyType String      
    }

    Write-Host -ForegroundColor Green "Mapped Successfully!"
    $body = @"
    {
        `"time_changed`": `"$(Get-Date -UFormat "%m/%d/%Y %R")`",
        `"type_of_change`":  `"Mapped Drive`",
        `"changed_user`":  `"$($wpflstbusers.selecteditem)`",
        `"addition_details`":  `"Mapped drive $driveletter to $path on $global:mappingcomputer`",
        `"ipAddress`":  `"`",
        `"computer`":  `"$global:hostname`",
        `"user`":  `"$global:whoami`"
    }  
"@

    Add-Content -Path $LogsPath -Value "`n$body`n"

})
#endregion

#region LAPS password -- Get either new LAPS (8/2023; LAPS tab in AD) or old LAPS (AD-schema ms-Mcs-AdmPwd)
$WPFbtnLAPSSearch.add_Click({
    $WPFtxtLAPSPassword.text = ''    
    
    $compname = $WPFtxtLAPSComputerName.Text
    $lapspw = (Get-LapsADPassword -AsPlainText -Identity $compname).password

    if ($null -eq $lapspw -or $lapspw -eq ''){
        $lapspw = Get-ADComputer $compname -Properties 'ms-Mcs-AdmPwd'
        $lapspw = $lapspw.'ms-Mcs-AdmPwd'
    }

    $WPFtxtLAPSPassword.Text = $lapspw    
})

$WPFbtnCopyLAPSPW.add_Click({
    Set-Clipboard -Value $WPFtxtLAPSPassword.Text
})

#endregion
poplockedusers

#Gets the current app theme for windows and applies either dark mode or light mode based on that
if ((Get-ItemProperty HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize | Select-Object AppsUseLightTheme).AppsUseLightTheme -eq 0){
    changeTheme
}

#Sets the LOG file directory, or creates one if non-existent
$LogsPath = "C:\logs\AAAA.log"
if (-not (Test-Path $LogsPath)){
    $LogsPath = (New-Item $LogsPath -Force).FullName
}


#[Console.Window]::ShowWindow($global:consolePtr, 0) | Out-Null
#Clear-Host
$Form.ShowDialog() | Out-Null
