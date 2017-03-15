[cmdletbinding()]
param(
    [Parameter(
        Mandatory = $false
    )]
    [Switch]
    $ConfigScript = $false,
    [Parameter(
        Mandatory = $false
    )]
    [ValidateSet('XML','JSON')]
    [String]
    $ImportAs = 'XML'
)

#Setup paths
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$inputDir   = "$scriptPath\Input"
$outputDir  = "$scriptPath\Output"

#Set the correct file extension
Switch ($ImportAs) {

    'JSON' {

        $importExtension = '.json'

    }

    'XML' {

        $importExtension = '.xml'

    }

}

#Set config file correctly
$configFile = "$inputDir\config$($importExtension)"

function Import-Config { #Begin function Import-Config
    [cmdletbinding()]
    param()

    Switch ($ImportAs) {

        'JSON' {

            $script:configData = Get-Content -Path $configFile | ConvertFrom-Json

        }

        'XML' {

            $script:configData  = Import-Clixml -Path $configFile

        }

    }   

} #End function Import-Config

function Export-Config { #Begin function Export-Config
    [cmdletbinding()]
    param(
        [Parameter(
            Mandatory
        )]
        [ValidateSet('JSON','XML')]
        $ExportAs
    )

    #If there's a file, we're gonna back it up!
    if (Test-Path -Path $configFile) {

        $backup = $true       

    }

    Switch ($ExportAs) { #Begin config type switch

        'JSON' {

            if ($backup) {
                
                Get-Content $configFile | Out-File -FilePath $configFile.Replace('.json','.json.bak')

                Write-Verbose "Backed up existing configuration to $($configFile.Replace('.json','.json.bak'))!"
                Write-Verbose ""

            }
            
            $ConfigurationOptions | ConvertTo-Json | Out-File -FilePath $configFile            

        }

        'XML' {

            if ($backup) {

                Get-Content $configFile | Out-File -FilePath $configFile.Replace('.xml','.xml.bak')

                Write-Verbose "Backed up existing configuration to $($configFile.Replace('.xml','.xml.bak'))!"
                Write-Verbose ""

            }            

            $ConfigurationOptions | Export-Clixml -Path $configFile

        }

    } #End config type switch

} #End function Export-Config

function Invoke-UserAction { #Begin function Invoke-UserAction
    [cmdletbinding()]
    param(
        [Parameter(
            Mandatory,
            ValueFromPipeline
        )]
        $usersToProcess
    )

    Begin { #Begin begin block for Invoke-UserAction

        #Create array to store results in
        [System.Collections.ArrayList]$processedArray = @()

        Write-Verbose "User processing started!"
        Write-Verbose ""

    } #End begin block for Invoke-UserAction

    Process { #Begin process block for function Invoke-UserAction

        foreach ($user in $usersToProcess) { #Begin user foreach loop
        
            #Set variables to null so they are not set by the last iteration
            $lastLogonDays = $null
            $userAction    = $null
            
            $notes         = 'N/A'

            #Some error handling for getting the last logon days
            Try {

                #Set value based on calculation using the LastLogon value of the user
                $lastLogonDays = ((Get-Date) - $user.LastLogon).Days 

            }

            Catch {
            
                #Capture message into variable $errorMessage, and set other variables accordingly
                $errorMessage  = $_.Exception.Message
                $lastLogonDays = $null
                $notes         = $errorMessage 

                Write-Warning "Issue encountered while calculating last logon days [$errorMessage]"   
                Write-Warning ""              

            }

            Write-Verbose "Checking on [$($user.DisplayName)], who last logged on [$lastLogonDays] days ago..." 
            Write-Verbose ""

            #Switch statement to switch out the value of $lastLogonDays
            Switch ($lastLogonDays) { #Begin action switch

                #This expression compares the value of $lastLogondays to the script scoped variable for warning days, set with the configuration data file
                {$_ -lt $script:configData.DisableDays -and $_ -ge $script:configData.WarnDays} { #Begin actions for warning

                    $userAction = 'Warn'

                    Write-Verbose "Warning, [$($user.DisplayName)] will be disabled in [$($script:configData.DisableDays - $lastLogonDays)] days!"
                    Write-Verbose ""

                    Break

                } #End actions for warning

                #This expression compares the value of $lastLogondays to the script scoped variable for disable days, set with the configuration data file
                {$_ -ge $script:configData.DisableDays} { #Begin actions for disable

                    $userAction = 'Disable'

                    Write-Verbose "[$($user.DisplayName)] is going to be disabled, and is [$($lastLogonDays - $script:ConfigData.DisableDays)] days past the threshold!"
                    Write-Verbose ""

                    Break

                } #End actions for disable

                {$_ -eq $null} { #Begin actions for a null value

                    $userAction = 'Error'

                    Write-Verbose "Something went wrong, no value specified for last logon days!"
                    Write-Verbose ""

                    Break

                } #End actions for a null value

                #Adding a default to catch other values
                default { #Begin default actions

                    $userAction = 'None'                    
                    Write-Verbose "$($user.DisplayName) is good to go, they last logged on [$($lastLogonDays)] days ago!"
                    Write-Verbose ""

                } #Begin default actions

            } #End action switch

            #Create object to store in array
            $processedObject = [PSCustomObject]@{
                
                DisplayName   = $user.DisplayName
                UserName      = $user.UserName
                OU            = $user.OU
                LastLogon     = $user.LastLogon
                LastLogonDays = $lastLogonDays
                Action        = $userAction                
                Notes         = $notes                

            }

            #Add object to array of processed users
            $processedArray.Add($processedObject) | Out-Null

        } #End user foreach loop

    } #End process block for function Invoke-UserAction

    End { #Begin end block for Invoke-UserAction

        Write-Verbose "User processing ended!"
        Write-Verbose ""
        
        #Return array
        Return $processedArray

    } #End end block for Invoke-UserAction

} #End function Invoke-UserAction

function Invoke-ConfigurationGeneration { #Begin function Invoke-ConfigurationGeneration
    [cmdletbinding()]
    param(
        [Parameter(
                Mandatory = $false
        )]
        [ValidateSet('XML','JSON')]
        [String]
        $ExportAs = $ImportAs,
        [Parameter(
            Mandatory = $false
        )]
        $ConfigurationOptions
    )

    if (!$configurationOptions) { #Actions if we don't pass in any options to the function
        
        #The OU list will be an array
        [System.Collections.ArrayList]$ouList = @()

        #These variables will be used to evaluate last logon dates of users
        [int]$warnDays    = 23
        [int]$disableDays = 30

        #Add some fake OUs for testing purposes
        $ouList.Add('OU=Marketing,DC=FakeDomain,DC=COM') | Out-Null
        $ouList.Add('OU=Sales,DC=FakeDomain,DC=COM')     | Out-Null

        #Create a custom object to store things in
        $configurationOptions = [PSCustomObject]@{

            WarnDays    = $warnDays
            DisableDays = $disableDays
            OUList      = $ouList

        }

        #Handle different types 
        #Export the object we created as the current configuration
        Export-Config -ExportAs $ExportAs

        Write-Verbose "Exporting generated configuration file to [$configFile]!"

    } else { #End actions for no options passed in, begin actions for if they are
        
        Export-Config -ExportAs $ExportAs

        Write-Verbose "Exporting passed in options as configuration file to [$configFile]!"

    } #End if for options passed into function

} #End function Invoke-ConfigurationGeneration

function Invoke-UserDiscovery { #Begin function Invoke-UserDiscovery
    [cmdletbinding()]
    param()

    #Create empty arrayList object
    [System.Collections.ArrayList]$userList = @()

    #Create users and add them to array
    $testUser2 = [PSCustomObject]@{

        DisplayName   = 'Mike Jones'
        UserName      = 'jonesm'
        LastLogon     = (Get-Date).AddDays(-35) 
        OU            = Get-Random -inputObject $script:configData.OUList

    }

    $testUser1 = [PSCustomObject]@{

        DisplayName   = 'John Doe'
        UserName      = 'doej'
        LastLogon     = (Get-Date).AddDays(-24) 
        OU            = Get-Random -inputObject $script:configData.OUList

    }    

    $testUser3 = [PSCustomObject]@{

        DisplayName   = 'Jim Doe'
        UserName      = 'doeji'
        LastLogon     = (Get-Date).AddDays(-10) 
        OU            = Get-Random -inputObject $script:configData.OUList

    }  
    

    $testUser4 = [PSCustomObject]@{

        DisplayName   = 'This WontWork'
        UserName      = 'wontworkt'        
        LastLogon     = $null 
        OU            = Get-Random -inputObject $script:configData.OUList

    }
    
    $testUser5 = [PSCustomObject]@{

        DisplayName   = 'This AlsoWontWork'
        UserName      = 'alsowontworkt'        
        LastLogon     = 'this many!'
        OU            = Get-Random -inputObject $script:configData.OUList

    }        

    $testUser6 = [PSCustomObject]@{

        DisplayName   = 'Sally Smith'
        UserName      = 'smiths'
        LastLogon     = (Get-Date).AddDays(-30) 
        OU            = Get-Random -inputObject $script:configData.OUList

    }  

    #Add users to arraylist
    $userList.Add($testUser1) | Out-Null
    $userList.Add($testUser2) | Out-Null
    $userList.Add($testUser3) | Out-Null
    $userList.Add($testUser4) | Out-Null
    $userList.Add($testUser5) | Out-Null
    $userList.Add($testUser6) | Out-Null

    #Return list
    Return $userList

} #End function Invoke-UserDiscovery

function Invoke-GUI { #Begin function Invoke-GUI
    [cmdletbinding()]
    Param()

    #We technically don't need these, but they may come in handy later if you want to pop up message boxes, etc
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

    #Input XAML here
    $inputXML = @"
<Window x:Class="psguiconfig.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:psguiconfig"
        mc:Ignorable="d"
        ResizeMode="NoResize"
        Title="Script Configuration" Height="281.26" Width="509.864"        
        >
    <Grid>
        <Label x:Name="lblWarningMin" Content="Warning Days" HorizontalAlignment="Left" Height="29" Margin="10,6,0,0" VerticalAlignment="Top" Width="86"/>
        <TextBox x:Name="txtBoxWarnLow" HorizontalAlignment="Left" Height="20" Margin="96,6,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="27"/>
        <Label x:Name="lblDisableMin" Content="Disable Days" HorizontalAlignment="Left" Height="29" Margin="134,6,0,0" VerticalAlignment="Top" Width="86"/>
        <TextBox x:Name="txtBoxDisableLow" HorizontalAlignment="Left" Height="20" Margin="220,6,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="27"/>
        <TextBox x:Name="txtBoxOUList" HorizontalAlignment="Left" Height="153" Margin="10,54,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="479" AcceptsReturn="True" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
        <Label x:Name="lblOUs" Content="OUs To Scan" HorizontalAlignment="Left" Height="26" Margin="10,28,0,0" VerticalAlignment="Top" Width="80"/>
        <Button x:Name="btnExceptions" Content="Exceptions" HorizontalAlignment="Left" Height="43" Margin="252,6,0,0" VerticalAlignment="Top" Width="237"/>
        <Button x:Name="btnEdit" Content="Edit" HorizontalAlignment="Left" Height="29" Margin="10,212,0,0" VerticalAlignment="Top" Width="66"/>
        <Button x:Name="btnSave" Content="Save" HorizontalAlignment="Left" Height="29" Margin="423,212,0,0" VerticalAlignment="Top" Width="66"/>
    </Grid>
</Window>
"@  

    [xml]$XAML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window' 
    
    #Read XAML 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
    try {
    
        $Form=[Windows.Markup.XamlReader]::Load( $reader )
        
    }

    catch {
    
        Write-Error "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
        
    }
 
    #Create variables to control form elements as objects in PowerShell
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
        
    } 
    
    #Setup the form    
    function Invoke-FormSetup { #Begin function Invoke-FormSetup

        #Here we set the default states of the objects that represent the buttons/fields
        $WPFbtnEdit.IsEnabled          = $true
        $WPFbtnSave.IsEnabled          = $false
        $WPFtxtBoxWarnLow.IsEnabled    = $false
        $WPFtxtBoxDisableLow.IsEnabled = $false
        $WPFtxtBoxOUList.IsEnabled     = $false
        $WPFbtnExceptions.IsEnabled    = $false

        #We will use the current values we imported from the script scoped variable configData
        $WPFtxtBoxWarnLow.Text    = $script:configData.WarnDays
        $WPFtxtBoxDisableLow.Text = $script:configData.DisableDays
        $WPFtxtBoxOUList.Text     = $script:configData.OUList | Out-String

    } #End function Invoke-FormSetup

    function Invoke-FormSaveData { #Begin function Invoke-FormSaveData

        #This function will perform the action to save the form data
        
        #We setup the variables based on the current values of the form
        $warnDays     = [int]$WPFtxtBoxWarnLow.Text 
        $disableDays  = [int]$WPFtxtBoxDisableLow.Text
        $ouList       = ($WPFtxtBoxOUList.Text | Out-String).Trim() -split '[\r\n]' | Where-Object {$_ -ne ''}

        #This object will contain the current configuration we would like to export
        $configurationOptions = [PSCustomObject]@{

            WarnDays    = $warnDays
            DisableDays = $disableDays
            OUList      = $ouList

        }
        
        #We then pass the configuration to the function we created earlier that will export the options we pass in
        Invoke-ConfigurationGeneration -configurationOptions $configurationOptions

        #Then we re-import the config file after it is exported via the function above        
        Import-Config 

        #Finally we revert the GUI to the original state, which will also reflect the lastest configuration that we just exported
        Invoke-FormSetup

    } #End function Invoke-FormSaveData

    #Now we perform actions using the functions we created, as well as code that runs when buttons are clicked

    #Run form setup on launch
    Invoke-FormSetup

    #Button actions
    $WPFbtnEdit.Add_Click{ #Begin edit button actions

        #This will 'open up' the form and allow fields to be edited
        $WPFbtnExceptions.IsEnabled    = $true
        $WPFbtnSave.IsEnabled          = $true
        $WPFtxtBoxWarnLow.IsEnabled    = $true
        $WPFtxtBoxDisableLow.IsEnabled = $true
        $WPFtxtBoxOUList.IsEnabled     = $true
        $WPFbtnExceptions.IsEnabled    = $true

    } #End edit button actions

    $WPFbtnSave.Add_Click{ #Begin save button actions

        #The save button calls the Invoke-FormSaveData function
        Invoke-FormSaveData

    } #End save button actions

    #Show the form
    $form.showDialog() | Out-Null

} #End function Invoke-GUI

#Check for config, generate if it doesn't exist
if (!(Test-Path -Path $configFile)) { 

    Write-Verbose "Configuration file does not exist, creating!" 
    
    #Call our function to generate the file
    Invoke-ConfigurationGeneration
    
    Import-Config     

} else {

    #Import file since it exists
    Import-Config 

}

#Script logic
if ($configScript) { #Begin if to see if $configScript is set to true

    #If it's true, run this function to launch the GUI
    Invoke-GUI
    
} else { #Begin if/else for script exeuction (non-config)

    #Simple example for using the OUList defined in the config file
    ForEach ($ou in $script:configData.OUList) { #Begin foreach loop for OU actions

        Write-Verbose "Performing action on [$ou]!"

    } #End foreach loop for OU actions

    #Create some test users
    $userList = Invoke-UserDiscovery

    #Take actions on each user and store results in $processedUsers
    $processedUsers = $userList | Invoke-UserAction
    
    #Create file name for data export
    $outputFileName = ("$outputDir\processedUsers_{0:MMddyy_HHmm}.csv" -f (Get-Date))
    
    #Export processed users various data types
    $processedUsers | Export-Csv -Path $outputFileName -NoTypeInformation
    $processedUsers | Export-Clixml -Path ($outputFileName -replace 'csv','xml')

    Write-Verbose "File exported to [$outputDir]!"

    #Take a look at the array
    return $processedUsers 

} #End if/else for script actions (non-config)