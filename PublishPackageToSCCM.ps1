<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    PublishPackageToSCCM
.SYNOPSIS
    GUI based application. Publish package to SCCM, create AD Groups and distribute software.
.DESCRIPTION
    Use variables section to change main settings. 
Please make sure that Remote Server Administration Toolkit is installed and SCCM Console is installed on the machine before using this application.
Application creates two device collections for install and uninstall as Required, Hidden and one user collection as Available install.
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '800,700'
$Form.text                       = "Publish Package to SCCM"
$Form.TopMost                    = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "PackageName"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.Anchor                   = 'top,right,left'
$Label1.location                 = New-Object System.Drawing.Point(20,20)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Publisher"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.Anchor                   = 'top,right,left'
$Label2.location                 = New-Object System.Drawing.Point(20,50)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "ApplicationName"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.Anchor                   = 'top,right,left'
$Label3.location                 = New-Object System.Drawing.Point(20,80)
$Label3.Font                     = 'Microsoft Sans Serif,10'

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Version"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.Anchor                   = 'top,right,left'
$Label4.location                 = New-Object System.Drawing.Point(20,110)
$Label4.Font                     = 'Microsoft Sans Serif,10'

$Groupbox1                       = New-Object system.Windows.Forms.Groupbox
$Groupbox1.height                = 50
$Groupbox1.width                 = 360
$Groupbox1.text                  = "Deployment Type"
$Groupbox1.location              = New-Object System.Drawing.Point(20,140)

$tbPackageName                   = New-Object system.Windows.Forms.TextBox
$tbPackageName.multiline         = $false
$tbPackageName.width             = 230
$tbPackageName.height            = 20
$tbPackageName.Anchor            = 'top,right,left'
$tbPackageName.location          = New-Object System.Drawing.Point(150,15)
$tbPackageName.Font              = 'Microsoft Sans Serif,10'

$tbPublisher                     = New-Object system.Windows.Forms.TextBox
$tbPublisher.multiline           = $false
$tbPublisher.width               = 230
$tbPublisher.height              = 20
$tbPublisher.Anchor              = 'top,right,left'
$tbPublisher.location            = New-Object System.Drawing.Point(150,45)
$tbPublisher.Font                = 'Microsoft Sans Serif,10'

$tbApplicationName               = New-Object system.Windows.Forms.TextBox
$tbApplicationName.multiline     = $false
$tbApplicationName.width         = 230
$tbApplicationName.height        = 20
$tbApplicationName.Anchor        = 'top,right,left'
$tbApplicationName.location      = New-Object System.Drawing.Point(150,75)
$tbApplicationName.Font          = 'Microsoft Sans Serif,10'

$tbVersion                       = New-Object system.Windows.Forms.TextBox
$tbVersion.multiline             = $false
$tbVersion.width                 = 230
$tbVersion.height                = 20
$tbVersion.Anchor                = 'top,right,left'
$tbVersion.location              = New-Object System.Drawing.Point(150,105)
$tbVersion.Font                  = 'Microsoft Sans Serif,10'

$rbMSI                           = New-Object system.Windows.Forms.RadioButton
$rbMSI.text                      = "MSI"
$rbMSI.AutoSize                  = $true
$rbMSI.width                     = 104
$rbMSI.height                    = 20
$rbMSI.location                  = New-Object System.Drawing.Point(10,20)
$rbMSI.Font                      = 'Microsoft Sans Serif,10'

$rbPSAppDeploy                   = New-Object system.Windows.Forms.RadioButton
$rbPSAppDeploy.text              = "PSAppDeploy"
$rbPSAppDeploy.AutoSize          = $true
$rbPSAppDeploy.width             = 104
$rbPSAppDeploy.height            = 20
$rbPSAppDeploy.location          = New-Object System.Drawing.Point(70,20)
$rbPSAppDeploy.Font              = 'Microsoft Sans Serif,10'

$rbBAT                           = New-Object system.Windows.Forms.RadioButton
$rbBAT.text                      = "Install and Uninstall BAT"
$rbBAT.AutoSize                  = $true
$rbBAT.width                     = 104
$rbBAT.height                    = 20
$rbBAT.location                  = New-Object System.Drawing.Point(180,20)
$rbBAT.Font                      = 'Microsoft Sans Serif,10'

$gbMSI                           = New-Object system.Windows.Forms.Groupbox
$gbMSI.height                    = 50
$gbMSI.width                     = 360
$gbMSI.Anchor                    = 'top,right,left'
$gbMSI.text                      = "MSI"
$gbMSI.location                  = New-Object System.Drawing.Point(20,200)

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "MSI Name"
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(10,20)
$Label5.Font                     = 'Microsoft Sans Serif,10'

$tbMSIName                       = New-Object system.Windows.Forms.TextBox
$tbMSIName.multiline             = $false
$tbMSIName.width                 = 210
$tbMSIName.height                = 20
$tbMSIName.Anchor                = 'top,right,left'
$tbMSIName.location              = New-Object System.Drawing.Point(130,15)
$tbMSIName.Font                  = 'Microsoft Sans Serif,10'

$gbActions                       = New-Object system.Windows.Forms.Groupbox
$gbActions.height                = 370
$gbActions.width                 = 370
$gbActions.Anchor                = 'top,right'
$gbActions.text                  = "Execute Actions "
$gbActions.location              = New-Object System.Drawing.Point(420,10)

$cbCreatePackage                 = New-Object system.Windows.Forms.CheckBox
$cbCreatePackage.text            = "Create Package"
$cbCreatePackage.AutoSize        = $true
$cbCreatePackage.width           = 95
$cbCreatePackage.height          = 20
$cbCreatePackage.location        = New-Object System.Drawing.Point(10,20)
$cbCreatePackage.Font            = 'Microsoft Sans Serif,10'

$cbCreateDeploymentType          = New-Object system.Windows.Forms.CheckBox
$cbCreateDeploymentType.text     = "Create Deployment Type"
$cbCreateDeploymentType.AutoSize  = $true
$cbCreateDeploymentType.width    = 95
$cbCreateDeploymentType.height   = 20
$cbCreateDeploymentType.location  = New-Object System.Drawing.Point(10,40)
$cbCreateDeploymentType.Font     = 'Microsoft Sans Serif,10'

$cbCreateDeviceCollections       = New-Object system.Windows.Forms.CheckBox
$cbCreateDeviceCollections.text  = "Create Device Collections"
$cbCreateDeviceCollections.AutoSize  = $true
$cbCreateDeviceCollections.width  = 95
$cbCreateDeviceCollections.height  = 20
$cbCreateDeviceCollections.location  = New-Object System.Drawing.Point(10,80)
$cbCreateDeviceCollections.Font  = 'Microsoft Sans Serif,10'

$cbCreateUserCollection          = New-Object system.Windows.Forms.CheckBox
$cbCreateUserCollection.text     = "Create User Collection"
$cbCreateUserCollection.AutoSize  = $true
$cbCreateUserCollection.width    = 95
$cbCreateUserCollection.height   = 20
$cbCreateUserCollection.location  = New-Object System.Drawing.Point(10,100)
$cbCreateUserCollection.Font     = 'Microsoft Sans Serif,10'

$btnStart                        = New-Object system.Windows.Forms.Button
$btnStart.text                   = "Start Executing Actions"
$btnStart.width                  = 180
$btnStart.height                 = 30
$btnStart.location               = New-Object System.Drawing.Point(20,350)
$btnStart.Font                   = 'Microsoft Sans Serif,10'

$errPackageName                  = New-Object system.Windows.Forms.Label
$errPackageName.text             = "*"
$errPackageName.AutoSize         = $true
$errPackageName.visible          = $false
$errPackageName.width            = 25
$errPackageName.height           = 10
$errPackageName.Anchor           = 'top,right'
$errPackageName.location         = New-Object System.Drawing.Point(390,15)
$errPackageName.Font             = 'Microsoft Sans Serif,10'
$errPackageName.ForeColor        = "#ff0000"

$errPublisher                    = New-Object system.Windows.Forms.Label
$errPublisher.text               = "*"
$errPublisher.AutoSize           = $true
$errPublisher.visible            = $false
$errPublisher.width              = 25
$errPublisher.height             = 10
$errPublisher.Anchor             = 'top,right'
$errPublisher.location           = New-Object System.Drawing.Point(390,45)
$errPublisher.Font               = 'Microsoft Sans Serif,10'
$errPublisher.ForeColor          = "#ff0000"

$errApplicationName              = New-Object system.Windows.Forms.Label
$errApplicationName.text         = "*"
$errApplicationName.AutoSize     = $true
$errApplicationName.visible      = $false
$errApplicationName.width        = 25
$errApplicationName.height       = 10
$errApplicationName.Anchor       = 'top,right'
$errApplicationName.location     = New-Object System.Drawing.Point(390,75)
$errApplicationName.Font         = 'Microsoft Sans Serif,10'
$errApplicationName.ForeColor    = "#ff0000"

$errVersion                      = New-Object system.Windows.Forms.Label
$errVersion.text                 = "*"
$errVersion.AutoSize             = $true
$errVersion.visible              = $false
$errVersion.width                = 25
$errVersion.height               = 10
$errVersion.Anchor               = 'top,right'
$errVersion.location             = New-Object System.Drawing.Point(390,105)
$errVersion.Font                 = 'Microsoft Sans Serif,10'
$errVersion.ForeColor            = "#ff0000"

$errMSIName                      = New-Object system.Windows.Forms.Label
$errMSIName.text                 = "*"
$errMSIName.AutoSize             = $true
$errMSIName.visible              = $false
$errMSIName.width                = 25
$errMSIName.height               = 10
$errMSIName.location             = New-Object System.Drawing.Point(345,15)
$errMSIName.Font                 = 'Microsoft Sans Serif,10'
$errMSIName.ForeColor            = "#ff0000"

$cbDeployToDeviceCollectionsHiddenRequired   = New-Object system.Windows.Forms.CheckBox
$cbDeployToDeviceCollectionsHiddenRequired.text  = "Deploy To Device Collections (Required, Hidden)"
$cbDeployToDeviceCollectionsHiddenRequired.AutoSize  = $true
$cbDeployToDeviceCollectionsHiddenRequired.width  = 95
$cbDeployToDeviceCollectionsHiddenRequired.height  = 20
$cbDeployToDeviceCollectionsHiddenRequired.location  = New-Object System.Drawing.Point(10,120)
$cbDeployToDeviceCollectionsHiddenRequired.Font  = 'Microsoft Sans Serif,10'

$cbDeployToUserCollectionAvailable   = New-Object system.Windows.Forms.CheckBox
$cbDeployToUserCollectionAvailable.text  = "Deploy To User Collection (Available)"
$cbDeployToUserCollectionAvailable.AutoSize  = $true
$cbDeployToUserCollectionAvailable.width  = 95
$cbDeployToUserCollectionAvailable.height  = 20
$cbDeployToUserCollectionAvailable.location  = New-Object System.Drawing.Point(10,140)
$cbDeployToUserCollectionAvailable.Font  = 'Microsoft Sans Serif,10'

$cbRemoveDeploymentsAbove        = New-Object system.Windows.Forms.CheckBox
$cbRemoveDeploymentsAbove.text   = "Remove Device Collections Deployments"
$cbRemoveDeploymentsAbove.AutoSize  = $true
$cbRemoveDeploymentsAbove.width  = 95
$cbRemoveDeploymentsAbove.height  = 20
$cbRemoveDeploymentsAbove.location  = New-Object System.Drawing.Point(10,160)
$cbRemoveDeploymentsAbove.Font   = 'Microsoft Sans Serif,10'

$cbCreateADGroups                = New-Object system.Windows.Forms.CheckBox
$cbCreateADGroups.text           = "Create Device AD Groups (_I,_U)"
$cbCreateADGroups.AutoSize       = $true
$cbCreateADGroups.width          = 95
$cbCreateADGroups.height         = 20
$cbCreateADGroups.location       = New-Object System.Drawing.Point(10,180)
$cbCreateADGroups.Font           = 'Microsoft Sans Serif,10'

$cbCreateUserADGroup             = New-Object system.Windows.Forms.CheckBox
$cbCreateUserADGroup.text        = "Create User AD Group (_A)"
$cbCreateUserADGroup.AutoSize    = $true
$cbCreateUserADGroup.width       = 95
$cbCreateUserADGroup.height      = 20
$cbCreateUserADGroup.location    = New-Object System.Drawing.Point(10,200)
$cbCreateUserADGroup.Font        = 'Microsoft Sans Serif,10'

$cbAddADGroupsQueriesToDeviceCollections   = New-Object system.Windows.Forms.CheckBox
$cbAddADGroupsQueriesToDeviceCollections.text  = "Add AD Groups Queries to Device Collections"
$cbAddADGroupsQueriesToDeviceCollections.AutoSize  = $true
$cbAddADGroupsQueriesToDeviceCollections.width  = 95
$cbAddADGroupsQueriesToDeviceCollections.height  = 20
$cbAddADGroupsQueriesToDeviceCollections.location  = New-Object System.Drawing.Point(10,220)
$cbAddADGroupsQueriesToDeviceCollections.Font  = 'Microsoft Sans Serif,10'

$cbAddADGroupQueryToUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddADGroupQueryToUserCollection.text  = "Add AD Group Query To User Collection"
$cbAddADGroupQueryToUserCollection.AutoSize  = $true
$cbAddADGroupQueryToUserCollection.width  = 95
$cbAddADGroupQueryToUserCollection.height  = 20
$cbAddADGroupQueryToUserCollection.location  = New-Object System.Drawing.Point(10,240)
$cbAddADGroupQueryToUserCollection.Font  = 'Microsoft Sans Serif,10'

$cbAddTestMachineToDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddTestMachineToDeviceCollection.text  = "Add Test Machines To Device Collection"
$cbAddTestMachineToDeviceCollection.AutoSize  = $true
$cbAddTestMachineToDeviceCollection.width  = 95
$cbAddTestMachineToDeviceCollection.height  = 20
$cbAddTestMachineToDeviceCollection.location  = New-Object System.Drawing.Point(10,260)
$cbAddTestMachineToDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveTestMachineFromDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveTestMachineFromDeviceCollection.text  = "Remove Test Machines From Device Collection"
$cbRemoveTestMachineFromDeviceCollection.AutoSize  = $true
$cbRemoveTestMachineFromDeviceCollection.width  = 95
$cbRemoveTestMachineFromDeviceCollection.height  = 20
$cbRemoveTestMachineFromDeviceCollection.location  = New-Object System.Drawing.Point(10,280)
$cbRemoveTestMachineFromDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbAddTestUserToUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddTestUserToUserCollection.text  = "Add Test Users To User Collection"
$cbAddTestUserToUserCollection.AutoSize  = $true
$cbAddTestUserToUserCollection.width  = 95
$cbAddTestUserToUserCollection.height  = 20
$cbAddTestUserToUserCollection.location  = New-Object System.Drawing.Point(10,300)
$cbAddTestUserToUserCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveTestUserFromUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveTestUserFromUserCollection.text  = "Remove Test Users From User Collection"
$cbRemoveTestUserFromUserCollection.AutoSize  = $true
$cbRemoveTestUserFromUserCollection.width  = 95
$cbRemoveTestUserFromUserCollection.height  = 20
$cbRemoveTestUserFromUserCollection.location  = New-Object System.Drawing.Point(10,320)
$cbRemoveTestUserFromUserCollection.Font  = 'Microsoft Sans Serif,10'

$Label6                          = New-Object system.Windows.Forms.Label
$Label6.text                     = "Test Machines (;)"
$Label6.AutoSize                 = $true
$Label6.width                    = 25
$Label6.height                   = 10
$Label6.location                 = New-Object System.Drawing.Point(20,260)
$Label6.Font                     = 'Microsoft Sans Serif,10'

$tbTestMachine                   = New-Object system.Windows.Forms.TextBox
$tbTestMachine.multiline         = $false
$tbTestMachine.width             = 230
$tbTestMachine.height            = 20
$tbTestMachine.Anchor            = 'top,right,left'
$tbTestMachine.location          = New-Object System.Drawing.Point(150,255)
$tbTestMachine.Font              = 'Microsoft Sans Serif,10'

$Label7                          = New-Object system.Windows.Forms.Label
$Label7.text                     = "Test Users (;)"
$Label7.AutoSize                 = $true
$Label7.width                    = 25
$Label7.height                   = 10
$Label7.location                 = New-Object System.Drawing.Point(21,285)
$Label7.Font                     = 'Microsoft Sans Serif,10'

$tbTestUser                      = New-Object system.Windows.Forms.TextBox
$tbTestUser.multiline            = $false
$tbTestUser.width                = 230
$tbTestUser.height               = 20
$tbTestUser.Anchor               = 'top,right,left'
$tbTestUser.location             = New-Object System.Drawing.Point(150,280)
$tbTestUser.Font                 = 'Microsoft Sans Serif,10'

$lbLog                           = New-Object system.Windows.Forms.ListBox
$lbLog.text                      = "Log"
$lbLog.width                     = 770
$lbLog.height                    = 130
$lbLog.Anchor                    = 'top,right,bottom,left'
$lbLog.location                  = New-Object System.Drawing.Point(20,400)

$cbDistributeContent             = New-Object system.Windows.Forms.CheckBox
$cbDistributeContent.text        = "Distribute Content"
$cbDistributeContent.AutoSize    = $true
$cbDistributeContent.width       = 95
$cbDistributeContent.height      = 20
$cbDistributeContent.location    = New-Object System.Drawing.Point(10,60)
$cbDistributeContent.Font        = 'Microsoft Sans Serif,10'

$cbUpdateContent                 = New-Object system.Windows.Forms.CheckBox
$cbUpdateContent.text            = "Update Content On Distribution Points"
$cbUpdateContent.AutoSize        = $true
$cbUpdateContent.width           = 95
$cbUpdateContent.height          = 20
$cbUpdateContent.location        = New-Object System.Drawing.Point(10,340)
$cbUpdateContent.Font            = 'Microsoft Sans Serif,10'

$ddPresets                       = New-Object system.Windows.Forms.ComboBox
$ddPresets.text                  = "Presets"
$ddPresets.width                 = 180
$ddPresets.height                = 20
$ddPresets.Anchor                = 'top,right'
@('Not Set','Create Package','Pass to Regression Testing','Pass to Live','Add Test User Deployment','Add User Deployment To Live') | ForEach-Object {[void] $ddPresets.Items.Add($_)}
$ddPresets.location              = New-Object System.Drawing.Point(180,13)
$ddPresets.Font                  = 'Microsoft Sans Serif,10'

$ddApplications                  = New-Object system.Windows.Forms.ComboBox
$ddApplications.text             = "Click Get Applications for Selected Folder button to Populate this list"
$ddApplications.width            = 770
$ddApplications.height           = 20
$ddApplications.visible          = $true
$ddApplications.enabled          = $false
$ddApplications.Anchor           = 'right,bottom,left'
$ddApplications.location         = New-Object System.Drawing.Point(20,570)
$ddApplications.Font             = 'Microsoft Sans Serif,10'

$btnRefreshApplications          = New-Object system.Windows.Forms.Button
$btnRefreshApplications.text     = "Get Applications for Selected Folder"
$btnRefreshApplications.width    = 260
$btnRefreshApplications.height   = 30
$btnRefreshApplications.visible  = $true
$btnRefreshApplications.enabled  = $false
$btnRefreshApplications.Anchor   = 'bottom,left'
$btnRefreshApplications.location  = New-Object System.Drawing.Point(280,600)
$btnRefreshApplications.Font     = 'Microsoft Sans Serif,10'

$btnLoadApplication              = New-Object system.Windows.Forms.Button
$btnLoadApplication.text         = "Load Application"
$btnLoadApplication.width        = 200
$btnLoadApplication.height       = 30
$btnLoadApplication.visible      = $true
$btnLoadApplication.enabled      = $false
$btnLoadApplication.Anchor       = 'bottom,left'
$btnLoadApplication.location     = New-Object System.Drawing.Point(590,600)
$btnLoadApplication.Font         = 'Microsoft Sans Serif,10'

$ddApplicationFolders            = New-Object system.Windows.Forms.ComboBox
$ddApplicationFolders.text       = "Click Refresh Application Folders button to Populate this list"
$ddApplicationFolders.width      = 770
$ddApplicationFolders.height     = 20
$ddApplicationFolders.Anchor     = 'right,bottom,left'
$ddApplicationFolders.location   = New-Object System.Drawing.Point(20,540)
$ddApplicationFolders.Font       = 'Microsoft Sans Serif,10'

$btnRefreshApplicationFolders    = New-Object system.Windows.Forms.Button
$btnRefreshApplicationFolders.text  = "Refresh Application Folders"
$btnRefreshApplicationFolders.width  = 200
$btnRefreshApplicationFolders.height  = 30
$btnRefreshApplicationFolders.Anchor  = 'bottom,left'
$btnRefreshApplicationFolders.location  = New-Object System.Drawing.Point(20,600)
$btnRefreshApplicationFolders.Font  = 'Microsoft Sans Serif,10'

$btnRemoveApplication            = New-Object system.Windows.Forms.Button
$btnRemoveApplication.text       = "Remove Application and its Deployments"
$btnRemoveApplication.width      = 329
$btnRemoveApplication.height     = 30
$btnRemoveApplication.visible    = $true
$btnRemoveApplication.enabled    = $false
$btnRemoveApplication.Anchor     = 'bottom,left'
$btnRemoveApplication.location   = New-Object System.Drawing.Point(20,650)
$btnRemoveApplication.Font       = 'Microsoft Sans Serif,10'

$tbUnikeyRef                     = New-Object system.Windows.Forms.TextBox
$tbUnikeyRef.multiline           = $false
$tbUnikeyRef.text                = "PFA-"
$tbUnikeyRef.width               = 230
$tbUnikeyRef.height              = 20
$tbUnikeyRef.Anchor              = 'top,right,left'
$tbUnikeyRef.location            = New-Object System.Drawing.Point(150,305)
$tbUnikeyRef.Font                = 'Microsoft Sans Serif,10'

$lblUnikeyRef                    = New-Object system.Windows.Forms.Label
$lblUnikeyRef.text               = "Unikey Ref"
$lblUnikeyRef.AutoSize           = $true
$lblUnikeyRef.width              = 25
$lblUnikeyRef.height             = 10
$lblUnikeyRef.location           = New-Object System.Drawing.Point(20,310)
$lblUnikeyRef.Font               = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($Label1,$Label2,$Label3,$Label4,$Groupbox1,$tbPackageName,$tbPublisher,$tbApplicationName,$tbVersion,$gbMSI,$gbActions,$btnStart,$errPackageName,$errPublisher,$errApplicationName,$errVersion,$Label6,$tbTestMachine,$Label7,$tbTestUser,$lbLog,$ddApplications,$btnRefreshApplications,$btnLoadApplication,$ddApplicationFolders,$btnRefreshApplicationFolders,$btnRemoveApplication,$tbUnikeyRef,$lblUnikeyRef))
$Groupbox1.controls.AddRange(@($rbMSI,$rbPSAppDeploy,$rbBAT))
$gbMSI.controls.AddRange(@($Label5,$tbMSIName,$errMSIName))
$gbActions.controls.AddRange(@($cbCreatePackage,$cbCreateDeploymentType,$cbCreateDeviceCollections,$cbCreateUserCollection,$cbDeployToDeviceCollectionsHiddenRequired,$cbDeployToUserCollectionAvailable,$cbRemoveDeploymentsAbove,$cbCreateADGroups,$cbCreateUserADGroup,$cbAddADGroupsQueriesToDeviceCollections,$cbAddADGroupQueryToUserCollection,$cbAddTestMachineToDeviceCollection,$cbRemoveTestMachineFromDeviceCollection,$cbAddTestUserToUserCollection,$cbRemoveTestUserFromUserCollection,$cbDistributeContent,$cbUpdateContent,$ddPresets))

$btnStart.Add_Click({ btnStartClicked })
$cbRemoveDeploymentsAbove.Add_CheckedChanged({ cbRemoveDeploymentsAboveCheckedChanged })
$Form.Add_Shown({ FormShown })
$ddPresets.Add_SelectedValueChanged({ PresetsChanged })
$btnRefreshApplications.Add_Click({ RefreshApplicationsList })
$btnLoadApplication.Add_Click({ LoadApplication })
$ddApplicationFolders.Add_SelectedValueChanged({ ApplicationFolderSelected })
$btnRefreshApplicationFolders.Add_Click({ RefreshApplicationFoldersClicked })
$btnRemoveApplication.Add_Click({ RemoveLoadedApplicationClicked })

#### Defaults

$lbLog.HorizontalScrollbar =$true
$rbMSI.checked = $true

#### Variables

$tbTestUser.text = $env:UserName
$tbTestMachine.text = "TESTMACHINE"

$SiteCode = "S01" # Site code 
$SiteServer = "SiteServer" # SMS Provider machine name

$PackageRepository = "\\SCCM\Applications"
$ADGroupNamePrefix = "" # e.g. SCCM_App_
$LogFileLocation = "$($env:TEMP)\PublishPackageToSCCM_$((get-date).tostring('ddMMyyHHmmss')).log"
$NewPackageLocation = "" # default package folder inside SCCM
$DistributionPointGroups = @("All distribution points")
$InstallDeviceCollectionLocation = "" # default install device collection folder inside SCCM
$UninstallDeviceCollectionLocation = "" # default uninstall device collection folder inside SCCM
$LimitingDeviceCollectionName = "All Systems" # your limiting collection name, e.g. Windows 10 Machines, All Systems
$LimitingUserCollectionName = "All Users" # your limiting collection name, e.g. Win10 Users, All Users
$UserCollectionLocation = "" # default user collection folder inside SCCM
$ADOUPath = "OU=My,DC=Domainname,DC=com" # Active Directory path where applications are stored
$DomainPrefix = "Domainname" # your domain prefix to be used in WMI query below 
$ADGroupQuery = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SystemGroupName = '$DomainPrefix\\"
$ADGroupUserQuery = "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,SMS_R_USER.Name,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_User where SMS_R_User.UserGroupName = '$DomainPrefix\\"


##### test - mode
$testMode = $false
if($testMode){

$btnRemoveApplication.enabled = $true

$tbPackageName.text = ""
$tbPublisher.text = ""
$tbApplicationName.text = ""
$tbVersion.text = ""
$tbMSIName.text = ""

$LogFileLocation = "$($env:TEMP)\1111.log"
}


#### Functions


function RemoveDotMSIAtTheEnd($s){
    return $(if($s -like "*.msi"){$s.substring(0,$s.length-4)}else{$s})
}

#Get ProductCode from MSI
function GetProductCode($MSIFullPath){
    $retVal = ""
    try{
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    	$MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MSIFullPath, 0))
        $Query = "SELECT Value FROM Property WHERE Property = 'ProductCode'"
    	$View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
    	$View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
    	$Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
    	$retVal = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
    }
    catch{
        $retVal += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    }
    return $retVal
}

function ClearForm{
    ClearErrors
    $lbLog.Items.Clear
}

function ClearErrors{
    $errPackageName.visible = $false
    $errPublisher.visible = $false
    $errApplicationName.visible = $false
    $errVersion.visible = $false
    $errMSIName.visible = $false
}

#Write Log
function wl($message){
    try{
        $tmpmessage = "$(Get-Date -Format 'HH:mm:ss') $message"
        $lbLog.Items.Add($tmpmessage)
        $lbLog.SelectedIndex = $lbLog.Items.Count - 1;
        $Form.refresh()
        $tmpmessage|Out-File -FilePath $LogFileLocation -Append
    }catch{
        ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    }
    return $message
}

function ConnectToSCCMAndAD{
    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
    }
    wl("Connected to SCCM")
    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ActiveDirectory) -eq $null) {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    wl("Connected to AD")
    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop
    }
    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" -ErrorAction Stop
}

function FormShown { ConnectToSCCMAndAD }

function RefreshApplicationFoldersClicked {
    wl("Refreshing application folders list...")
    $Instancekeys = (Get-WmiObject -Namespace "ROOT\SMS\Site_$Sitecode" -ComputerName $SiteServer `
         -Query "select DISTINCT ObjectPath from SMS_Applicationlatest where ModelName in (select InstanceKey from SMS_ObjectContainerItem where ObjectType=6000) ORDER BY ObjectPath").ObjectPath
    $ddApplicationFolders.Items.Clear
    $ddApplicationFolders.Items.Add("All Applications from all folders")
    foreach ($key in $Instancekeys){
        $ddApplicationFolders.Items.Add($key)
    }
    $ddApplicationFolders.SelectedItem = "All Applications from all folders" #$ddApplicationFolders.Items[0]
    wl("Populated $($Instancekeys.Count) folders")
    $btnRefreshApplications.enabled = $true
}

function RefreshApplicationsList { 
    wl("Refreshing applications list...")
    if($ddApplicationFolders.SelectedItem -eq "All Applications from all folders"){
        $Instancekeys = (Get-WmiObject -Namespace "ROOT\SMS\Site_$Sitecode" -ComputerName $SiteServer `
-Query "select LocalizedDisplayName from SMS_Applicationlatest ORDER BY LocalizedDisplayName").LocalizedDisplayName
    }else{
        $Instancekeys = (Get-WmiObject -Namespace "ROOT\SMS\Site_$Sitecode" -ComputerName $SiteServer `
-Query "select LocalizedDisplayName from SMS_Applicationlatest where ObjectPath='$($ddApplicationFolders.SelectedItem)' ORDER BY LocalizedDisplayName").LocalizedDisplayName
    }
    $ddApplications.Items.Clear
    foreach ($key in $Instancekeys){
        $ddApplications.Items.Add($key)
    }
    $ddApplications.SelectedItem = $ddApplications.Items[0]
    $ddApplications.enabled = $true
    wl("Populated $($Instancekeys.Count) applications")
    $btnRemoveApplication.enabled = $true
    $btnLoadApplication.enabled = $true
}

function LoadApplication { 
    try{
        $app = Get-CMApplication $($ddApplications.SelectedItem)
        $tbPackageName.text = $ddApplications.SelectedItem
        $tbPublisher.text = $(if($app.Manufacturer -eq ""){""}else{$app.Manufacturer})
        $tbVersion.text = $(if($app.SoftwareVersion -eq ""){""}else{$app.SoftwareVersion})
        $tmppub = $(if($tbPublisher.text -eq ""){"_"}else{$tbPublisher.text})
        $tmpver = $(if($app.SoftwareVersion -eq ""){"_"}else{$($tbVersion.text).replace(".","")})
        $tbApplicationName.text = $($tbPackageName.text).replace("$($tmppub)_","").replace("$($tmpver)_","").replace("EN_01_W10_F","").replace("EN_02_W10_F","").replace("EN_03_W10_F","").replace("EN_04_W10_F","").replace("EN_05_W10_F","").replace("_"," ").Trim()
        if($tbApplicationName.text -eq ""){$tbApplicationName.text = $tbPublisher.text}
        wl("Loaded an Application $($ddApplications.SelectedItem)")
    }
    catch{
        wl("Failed to load an Application $($ddApplications.SelectedItem)")
        wl($_.Exception.Message)
    }
    finally{
    }
}

function ShowMessageBoxWithError($errmsg){
    [System.Windows.Forms.MessageBox]::Show($errmsg,'Error','OK','Error')
    wl($errmsg)
    return $errmsg
}


function ApplicationFolderSelected {
    $ddApplications.visible = $true
    $btnRefreshApplications.visible = $true
}

function CheckApplicationExists($appname){
    $retval = $true
    try{
        #To avoid wildcards
        $appname = $appname.Replace("*","")
        $tmpObj = Get-CMApplication -Name "$appname"
        if($tmpObj -eq $null){
            $tmpstr = ShowMessageBoxWithError("Error: Application $appname wasn't found")
            $retval = $false
        }
    }
    catch{
        wl($_.Exception.Message)
        $retval = $false
    }
    return $retval
}




#### Presets


function PresetsChanged { 
	$ddPresets.SelectedItem
    switch($ddPresets.SelectedItem){
        "Not Set" {PresetNotSet; break}
        "Create Package" {PresetCreatePackage; break}
        "Pass to Regression Testing" {PresetPasstoRegressionTesting; break}
        "Pass to Live" {PresetPasstoLive; break}
        "Add Test User Deployment" {PresetAddTestUserDeployment; break}
        "Add User Deployment To Live" {PresetAddUserDeploymentToLive; break}
    }
}
function PresetNotSet { 
    $cbCreatePackage.checked = $false
    $cbCreateDeploymentType.checked = $false
    $cbDistributeContent.checked = $false
    $cbCreateDeviceCollections.checked = $false
    $cbCreateUserCollection.checked = $false
    $cbDeployToDeviceCollectionsHiddenRequired.checked = $false
    $cbDeployToUserCollectionAvailable.checked = $false
    $cbRemoveDeploymentsAbove.checked = $false
    $cbCreateADGroups.checked = $false
    $cbCreateUserADGroup.checked = $false
    $cbAddADGroupsQueriesToDeviceCollections.checked = $false
    $cbAddADGroupQueryToUserCollection.checked = $false
    $cbAddTestMachineToDeviceCollection.checked = $false
    $cbRemoveTestMachineFromDeviceCollection.checked = $false
    $cbAddTestUserToUserCollection.checked = $false
    $cbRemoveTestUserFromUserCollection.checked = $false
    $cbUpdateContent.checked = $false
}
function PresetCreatePackage { 
    $cbCreatePackage.checked = $true
    $cbCreateDeploymentType.checked = $true
    $cbDistributeContent.checked = $true
    $cbCreateDeviceCollections.checked = $true
    $cbCreateUserCollection.checked = $false
    $cbDeployToDeviceCollectionsHiddenRequired.checked = $true
    $cbDeployToUserCollectionAvailable.checked = $false
    $cbRemoveDeploymentsAbove.checked = $false
    $cbCreateADGroups.checked = $false
    $cbCreateUserADGroup.checked = $false
    $cbAddADGroupsQueriesToDeviceCollections.checked = $false
    $cbAddADGroupQueryToUserCollection.checked = $false
    $cbAddTestMachineToDeviceCollection.checked = $true
    $cbRemoveTestMachineFromDeviceCollection.checked = $false
    $cbAddTestUserToUserCollection.checked = $false
    $cbRemoveTestUserFromUserCollection.checked = $false
    $cbUpdateContent.checked = $false
}
function PresetPasstoRegressionTesting {
    $cbCreatePackage.checked = $false
    $cbCreateDeploymentType.checked = $false
    $cbDistributeContent.checked = $false
    $cbCreateDeviceCollections.checked = $false
    $cbCreateUserCollection.checked = $false
    $cbDeployToDeviceCollectionsHiddenRequired.checked = $false
    $cbDeployToUserCollectionAvailable.checked = $false
    $cbRemoveDeploymentsAbove.checked = $false
    $cbCreateADGroups.checked = $true
    $cbCreateUserADGroup.checked = $false
    $cbAddADGroupsQueriesToDeviceCollections.checked = $false
    $cbAddADGroupQueryToUserCollection.checked = $false
    $cbAddTestMachineToDeviceCollection.checked = $false
    $cbRemoveTestMachineFromDeviceCollection.checked = $true
    $cbAddTestUserToUserCollection.checked = $false
    $cbRemoveTestUserFromUserCollection.checked = $false
    $cbUpdateContent.checked = $false
}
function PresetPasstoLive { 
    $cbCreatePackage.checked = $false
    $cbCreateDeploymentType.checked = $false
    $cbDistributeContent.checked = $false
    $cbCreateDeviceCollections.checked = $false
    $cbCreateUserCollection.checked = $false
    $cbDeployToDeviceCollectionsHiddenRequired.checked = $false
    $cbDeployToUserCollectionAvailable.checked = $false
    $cbRemoveDeploymentsAbove.checked = $false
    $cbCreateADGroups.checked = $false
    $cbCreateUserADGroup.checked = $false
    $cbAddADGroupsQueriesToDeviceCollections.checked = $true
    $cbAddADGroupQueryToUserCollection.checked = $false
    $cbAddTestMachineToDeviceCollection.checked = $false
    $cbRemoveTestMachineFromDeviceCollection.checked = $false
    $cbAddTestUserToUserCollection.checked = $false
    $cbRemoveTestUserFromUserCollection.checked = $false
    $cbUpdateContent.checked = $false
}
function PresetAddTestUserDeployment { 
    $cbCreateUserCollection.checked = $true
    $cbDeployToUserCollectionAvailable.checked = $true
    $cbAddTestUserToUserCollection.checked = $true
}
function PresetAddUserDeploymentToLive {
    $cbCreateUserADGroup.checked = $true
    $cbAddADGroupQueryToUserCollection.checked = $true
    $cbRemoveTestUserFromUserCollection.checked = $true
}




#### Form Validation


function cbRemoveDeploymentsAboveCheckedChanged { 
    if($cbRemoveDeploymentsAbove.checked -eq $true){
        $cbDeployToUserCollectionAvailable.checked = $false
        $cbDeployToDeviceCollectionsHiddenRequired.checked = $false
    }
}

function ValidateForm{
    ClearErrors
	$retValidateForm = ""
    if(![System.IO.Directory]::Exists($defaultPackageLocation)){
        $errPackageName.visible = $true
        $retValidateForm += wl("Package Folder $defaultPackageLocation Doesn't Exist;")
    }
    if($tbPackageName.text -eq "" -Or $tbPackageName.text.length -gt 49){
        $errPackageName.visible = $true
        $retValidateForm += wl("PackageName is empty or over 49 characters long;")
    }
    if($tbPublisher.text -eq ""){
        $errPublisher.visible = $true
        $retValidateForm += wl("Publisher is empty;")
    }
    if($tbApplicationName.text -eq ""){
        $errApplicationName.visible = $true
        $retValidateForm += wl("ApplicationName is empty;")
    }
    if($tbVersion.text -eq ""){
        $errVersion.visible = $true
        $retValidateForm += wl("Version is empty;")
    }
    $cbCreatePackage.checked = $true
    $cbCreateDeploymentType.checked = $true
    $cbDistributeContent.checked = $true
    $cbCreateDeviceCollections.checked = $true
    $cbCreateUserCollection.checked = $false
    $cbDeployToDeviceCollectionsHiddenRequired.checked = $true
    $cbDeployToUserCollectionAvailable.checked = $false
    $cbRemoveDeploymentsAbove.checked = $false
    $cbCreateADGroups.checked = $false
    $cbCreateUserADGroup.checked = $false
    $cbAddADGroupsQueriesToDeviceCollections.checked = $false
    $cbAddADGroupQueryToUserCollection.checked = $false
    $cbAddTestMachineToDeviceCollection.checked = $true
    $cbRemoveTestMachineFromDeviceCollection.checked = $false
    $cbAddTestUserToUserCollection.checked = $false
    $cbRemoveTestUserFromUserCollection.checked = $false
    $cbUpdateContent.checked = $false
    if($tbMSIName.text -eq "" -And $rbMSI.checked -And $cbCreateDeploymentType.checked -And ($cbCreatePackage.checked -Or $cbCreateDeploymentType.checked -Or $cbDistributeContent.checked)){
        $errMSIName.visible = $true
        $retValidateForm += wl("MSIName is empty;")
    }
    $MSIName = RemoveDotMSIAtTheEnd($tbMSIName.text)
    if(![System.IO.File]::Exists("$defaultPackageLocation\$($MSIName).msi") -And $rbMSI.checked){
        $errMSIName.visible = $true
        $retValidateForm += wl("MSI doesn't exist $defaultPackageLocation\$($MSIName).msi")
    }
    
    return $retValidateForm
}



#### Remove Application


function RemoveLoadedApplicationClicked {
    $PackageName = $tbPackageName.text
    if($PackageName -eq ""){
        $PackageName=$ddApplications.SelectedItem
    }
    $UserResponse= [System.Windows.Forms.MessageBox]::Show("Are you sure you want to continue and remove application $($PackageName) from SCCM?" , "WARNING!!!" , 4)
    $returnerror = ""
    
    if ($UserResponse -eq "Yes"){
        wl("Remove Application: checking if application $PackageName exists")
        $ApplicationExists = CheckApplicationExists($PackageName)
        if ($ApplicationExists -eq $true){
            $UserResponse= [System.Windows.Forms.MessageBox]::Show("Do you want to export application to you Temp folder just in case something goes wrong?" , "WARNING!!!" , 4)
            if ($UserResponse -eq "Yes"){
                wl("Remove Application: Exporting application $($PackageName) to $($env:Temp)\$($PackageName).zip")
                wl("Remove Application: Export-CMApplication -Name ""$PackageName"" -Path ""$($env:Temp)\$($PackageName).zip"" -Force")
    			try{				
    				Export-CMApplication -Name "$PackageName" -Path "$($env:Temp)\$($PackageName).zip" -Force
    			}
    			catch{
    				$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    			}
            }
            if($returnerror -eq ""){
                wl("Remove Application: Enumirating Device collections...")
                $DeploymentCollections = @(Get-CMApplicationDeployment -Name "$PackageName"|Select -ExpandProperty CollectionName)
                wl("Remove Application: Found $($DeploymentCollections.Count) collections with deployments for this package")
                foreach($DeploymentCollection in $DeploymentCollections){
                    $uniqueAppInDeployment = $true
                    foreach($app in $(Get-CMApplicationDeployment -CollectionName $DeploymentCollection|Select -ExpandProperty ApplicationName)){
                        if($app -ne $PackageName){$uniqueAppInDeployment = $false}
                    }
                    if($uniqueAppInDeployment){
                        wl("Remove Application: Removing collection $DeploymentCollection as it doesn't have any other deployments")
                        wl("Remove Application: Remove-CMCollection -Name ""$DeploymentCollection"" -Force")
                        try{				
            				Remove-CMCollection -Name "$DeploymentCollection" -Force
            			}
            			catch{
    				        $returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            			}
                    }else{
                        wl("Remove Application: Leaving collection $DeploymentCollection behind as it has other deployments but removing the deployment")
                        wl("Remove Application: Remove-CMApplicationDeployment -Name ""$PackageName"" -CollectionName ""$DeploymentCollection"" -Force")
                        try{				
            				Remove-CMApplicationDeployment -Name "$PackageName" -CollectionName "$DeploymentCollection" -Force
            			}
            			catch{
    				        $returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            			}
                    }
                    if($returnerror -ne ""){break;}
                }
                if($returnerror -eq ""){
                    wl("Remove Application: Removing application from SCCM")
                    wl("Remove Application: Remove-CMApplication -Name ""$PackageName"" -Force")
                    try{
        				Remove-CMApplication -Name "$PackageName" -Force
                        wl("Remove Application: Application $PackageName was removed successfully")
        			}
        			catch{
    				    $returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        			}
                }
            }
        }
    }
}


#### Execute Actions


function btnStartClicked { 
    $returnerror = ""
    $PackageName = $tbPackageName.text
    $Publisher = $tbPublisher.text
    $ApplicationName = $tbApplicationName.text
    $Version = $tbVersion.text
    $MSIName = RemoveDotMSIAtTheEnd($tbMSIName.text)
    $testmachine = $($tbTestMachine.text).Split(";")
    $testuser = $($tbTestUser.text).Split(";")
    
    #calculated variables
    $InstallCollectionName = "$($PackageName)_I"
    $UninstallCollectionName = "$($PackageName)_U"
    $UserCollectionName = "$($PackageName)_A"
    $defaultPackageLocation = "$PackageRepository\$PackageName"
    $description = "$($tbUnikeyRef.text) $Publisher $ApplicationName $Version $defaultPackageLocation"
    $localizedName = "$Publisher $ApplicationName $Version"
    
    $returnerror += ValidateForm
    
    if($returnerror -eq ""){
        wl("Started Executing Actions")
    }
    
    #Create Package
    if($cbCreatePackage.checked -And $returnerror -eq ""){
        wl("Create Package: New-CMApplication -Name ""$PackageName"" -Description ""$description"" -Publisher ""$Publisher"" -SoftwareVersion ""$Version"" -LocalizedName ""$PackageName"" -LocalizedDescription ""$PackageName""")
        try{
			New-CMApplication -Name "$PackageName" -Description "$description" -Publisher "$Publisher" -SoftwareVersion "$Version" -LocalizedName "$PackageName" -LocalizedDescription "$PackageName"
			wl("Create Package: created application $PackageName")
		}
		catch{
		    $returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
		if($returnerror -eq ""){
		    if($NewPackageLocation -ne ""){
                wl("Create Package: Move-CMObject -InputObject `$(Get-CMApplication -Name ""$PackageName"") -FolderPath ""$($SiteCode):\Application\$NewPackageLocation""")
                try{
        			Move-CMObject -InputObject $(Get-CMApplication -Name "$PackageName") -FolderPath "$($SiteCode):\Application\$NewPackageLocation"
        			wl("Create Package: moved application to $NewPackageLocation")
        		}
        		catch{
        		    $returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        		}
		    }
		}
    }
    #Create Deployment Type
    if($cbCreateDeploymentType.checked -And $returnerror -eq ""){
        if($rbMSI.checked){
            wl("Create Deployment Type: getting ProductCode property for $defaultPackageLocation\$($MSIName).msi")
            $tmp = GetProductCode("$defaultPackageLocation\$($MSIName).msi")
            $ProductCode = ($tmp|Out-String).Trim()
            if($ProductCode -like "*error*"){
                $returnerror += $ProductCode
            }else{
                wl("Create Deployment Type: ProductCode=$ProductCode")
                $guidProductCode = [GUID]$ProductCode
            	$InstallCommand = "msiexec.exe /i $($MSIName).msi"
            	If([System.IO.File]::Exists("$defaultPackageLocation\$($PackageName).mst")){$InstallCommand += " TRANSFORMS=$($PackageName).mst"}
            	$InstallCommand += " /qn /l* C:\Windows\Logs\$($PackageName)_I.log"
            	wl("Create Deployment Type: InstallCommand=$InstallCommand")
            	$UninstallCommand = "msiexec.exe /x $ProductCode /qn /l* C:\Windows\Logs\$($PackageName)_U.log"
            	wl("Create Deployment Type: UninstallCommand=$UninstallCommand")
            	wl("Create Deployment Type: Add-CMMsiDeploymentType -DeploymentTypeName `"$PackageName`" -InstallCommand `"$InstallCommand`" -ApplicationName `"$PackageName`" -ProductCode $ProductCode -ContentLocation `"$defaultPackageLocation`" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand `"$UninstallCommand`" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment `"$description`" -Force")
                try{
                	Add-CMMsiDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "$InstallCommand" -ApplicationName "$PackageName" -ProductCode $ProductCode -ContentLocation "$defaultPackageLocation\$($MSIName).msi" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "$UninstallCommand" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description" -Force
                	wl("Create Deployment Type: created MSI deployment type")
                }catch{
                	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
                }
            }
        }
        if($rbPSAppDeploy.checked){
            wl("Create Deployment Type: Add-CMScriptDeploymentType -DeploymentTypeName ""$PackageName"" -InstallCommand ""Deploy-Application.exe -DeploymentType """"Install"""" -DeployMode """"Silent"""""" -ApplicationName ""$PackageName"" -AddDetectionClause `$(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName ""SOFTWARE\Freshfields\Installed\$PackageName"" -PropertyType String -ValueName ""Installed"" -Is64Bit -Existence) -ContentLocation ""$defaultPackageLocation"" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand ""Deploy-Application.exe -DeploymentType """"Uninstall"""" -DeployMode """"Silent"""""" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment ""$description""")
            try{
            	Add-CMScriptDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "Deploy-Application.exe -DeploymentType ""Install"" -DeployMode ""Silent""" -ApplicationName "$PackageName" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName "SOFTWARE\Freshfields\Installed\$PackageName" -PropertyType String -ValueName "Installed" -Is64Bit -Existence) -ContentLocation "$defaultPackageLocation" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "Deploy-Application.exe -DeploymentType ""Uninstall"" -DeployMode ""Silent""" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description"
            	wl("Create Deployment Type: created PSAppdeploy scripted deployment type")
            }catch{
            	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            }
        }
        if($rbBAT.checked){
            wl("Create Deployment Type: Add-CMScriptDeploymentType -DeploymentTypeName ""$PackageName"" -InstallCommand ""install.bat"" -ApplicationName ""$PackageName"" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName ""SOFTWARE\Freshfields\Installed\$PackageName"" -PropertyType String -ValueName ""Installed"" -Is64Bit -Existence) -ContentLocation ""$defaultPackageLocation"" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand ""uninstall.bat"" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment ""$description""")
            try{
            	Add-CMScriptDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "install.bat" -ApplicationName "$PackageName" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName "SOFTWARE\Freshfields\Installed\$PackageName" -PropertyType String -ValueName "Installed" -Is64Bit -Existence) -ContentLocation "$defaultPackageLocation" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "uninstall.bat" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description"
            	wl("Create Deployment Type: created BAT scripted deployment type")
            }catch{
            	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            }
        }
    }
    #Distribute Content
    if($cbDistributeContent.checked -And $returnerror -eq ""){
        wl("Distribute Content: Start-CMContentDistribution -ApplicationName ""$PackageName"" -DisableContentDependencyDetection -DistributionPointGroupName $DistributionPointGroups")
        try{
        	Start-CMContentDistribution -ApplicationName "$PackageName" -DisableContentDependencyDetection -DistributionPointGroupName $DistributionPointGroups
        	wl("Distribute Content: distributed content to $DistributionPointGroups")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Create Device Collections
    if($cbCreateDeviceCollections.checked -And $returnerror -eq ""){
        wl("Create Device Collections: New-CMDeviceCollection -Name ""$InstallCollectionName"" -LimitingCollectionName ""$LimitingDeviceCollectionName"" -RefreshType Continuous -Comment ""$description""")
        try{
        	New-CMDeviceCollection -Name "$InstallCollectionName" -LimitingCollectionName "$LimitingDeviceCollectionName" -RefreshType Continuous -Comment "$description"
        	wl("Create Device Collections: created Install Device Collection $InstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
        if($InstallDeviceCollectionLocation -ne ""){
            wl("Create Device Collections: Move-CMObject -InputObject `$(Get-CMDeviceCollection -Name ""$InstallCollectionName"") -FolderPath ""$($SiteCode):\DeviceCollection\$InstallDeviceCollectionLocation""")
            try{
            	Move-CMObject -InputObject $(Get-CMDeviceCollection -Name "$InstallCollectionName") -FolderPath "$($SiteCode):\DeviceCollection\$InstallDeviceCollectionLocation"
            	wl("Create Device Collections: moved it to $InstallDeviceCollectionLocation")
            }catch{
            	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            }
        }
        wl("Create Device Collections: New-CMDeviceCollection -Name ""$UninstallCollectionName"" -LimitingCollectionName ""$LimitingDeviceCollectionName"" -RefreshType Continuous -Comment ""$description""")
        try{
        	New-CMDeviceCollection -Name "$UninstallCollectionName" -LimitingCollectionName "$LimitingDeviceCollectionName" -RefreshType Continuous -Comment "$description"
        	wl("Create Device Collections: created Install Device Collection $UninstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
        if($UninstallDeviceCollectionLocation -ne ""){
            wl("Create Device Collections: Move-CMObject -InputObject `$(Get-CMDeviceCollection -Name ""$UninstallCollectionName"") -FolderPath ""$($SiteCode):\DeviceCollection\$UninstallDeviceCollectionLocation""")
            try{
            	Move-CMObject -InputObject $(Get-CMDeviceCollection -Name "$UninstallCollectionName") -FolderPath "$($SiteCode):\DeviceCollection\$UninstallDeviceCollectionLocation"
            	wl("Create Device Collections: moved it to $UninstallDeviceCollectionLocation")
            }catch{
            	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            }
        }
    }
    #Create User Collection
    if($cbCreateUserCollection.checked -And $returnerror -eq ""){
        wl("Create User Collection: New-CMUserCollection -Name ""$UserCollectionName"" -LimitingCollectionName ""$LimitingUserCollectionName"" -RefreshType Continuous")
        try{
        	New-CMUserCollection -Name "$UserCollectionName" -LimitingCollectionName "$LimitingUserCollectionName" -RefreshType Continuous
        	wl("Create User Collection: created Install User Collection $UserCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
        if($UserCollectionLocation -ne ""){
            wl("Create User Collection: Move-CMObject -InputObject $(Get-CMUserCollection -Name ""$UserCollectionName"") -FolderPath ""$($SiteCode):\UserCollection\$UserCollectionLocation""")
            try{
            	Move-CMObject -InputObject $(Get-CMUserCollection -Name "$UserCollectionName") -FolderPath "$($SiteCode):\UserCollection\$UserCollectionLocation"
        	    wl("Create User Collection: moved it to $UserCollectionLocation")
            }catch{
            	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            }
        }
    }
    #Deploy To Device Collections (Required, Hidden)
    if($cbDeployToDeviceCollectionsHiddenRequired.checked -And $returnerror -eq ""){
        wl("Deploy To Device Collections (Required, Hidden): New-CMApplicationDeployment -Name ""$PackageName"" -DeployAction Install -DeployPurpose Required -UserNotification HideAll -CollectionName ""$InstallCollectionName"" -Comment ""$description""")
        try{
        	New-CMApplicationDeployment -Name "$PackageName" -DeployAction Install -DeployPurpose Required -UserNotification HideAll -CollectionName "$InstallCollectionName" -Comment "$description"
        	wl("Deploy To Device Collections (Required, Hidden): assigned application for deployment to $InstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
        wl("Deploy To Device Collections (Required, Hidden): New-CMApplicationDeployment -Name ""$PackageName"" -DeployAction Uninstall -DeployPurpose Required -UserNotification HideAll -CollectionName ""$UninstallCollectionName"" -Comment ""$description""")
        try{
        	New-CMApplicationDeployment -Name "$PackageName" -DeployAction Uninstall -DeployPurpose Required -UserNotification HideAll -CollectionName "$UninstallCollectionName" -Comment "$description"
        	wl("Deploy To Device Collections (Required, Hidden): assigned application for deployment to $UninstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Deploy To User Collection (Available)
    if($cbDeployToUserCollectionAvailable.checked -And $returnerror -eq ""){
        wl("Deploy To User Collection (Available): New-CMApplicationDeployment -Name ""$PackageName"" -DeployAction Install -DeployPurpose Available -UserNotification DisplayAll -CollectionName ""$UserCollectionName"" -Comment ""$description""")
        try{
        	New-CMApplicationDeployment -Name "$PackageName" -DeployAction Install -DeployPurpose Available -UserNotification DisplayAll -CollectionName "$UserCollectionName" -Comment "$description"
        	wl("Deploy To User Collection (Available): assigned application for deployment to $UserCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Remove Device Collections Deployments
    if($cbRemoveDeploymentsAbove.checked -And $returnerror -eq ""){
        wl("Remove Device Collections Deployments: Remove-CMApplicationDeployment -Name ""$PackageName"" -Force -CollectionName ""$InstallCollectionName""")
        try{
        	Remove-CMApplicationDeployment -Name "$PackageName" -Force -CollectionName "$InstallCollectionName"
        	wl("Remove Device Collections Deployments: removed deployment for $InstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
        wl("Remove Device Collections Deployments: Remove-CMApplicationDeployment -Name ""$PackageName"" -Force -CollectionName ""$UninstallCollectionName""")
        try{
        	Remove-CMApplicationDeployment -Name "$PackageName" -Force -CollectionName "$UninstallCollectionName"
        	wl("Remove Device Collections Deployments: removed deployment for $UninstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Create Device AD Groups (_I,_U)
    if($cbCreateADGroups.checked -And $returnerror -eq ""){
        wl("Create Device AD Groups (_I,_U): New-ADGroup ""$ADGroupNamePrefix$($InstallCollectionName)"" -Path ""$ADOUPath"" -GroupCategory Security -GroupScope Global -Description ""$localizedName""")
		try{
			New-ADGroup "$ADGroupNamePrefix$($InstallCollectionName)" -Path "$ADOUPath" -GroupCategory Security -GroupScope Global -Description "$localizedName" `
-OtherAttributes @{info="This group deploys $($tbUnikeyRef.text) `"$PackageName`" to any Windows 10 devices specified - Please do not add users to the group. Devices should only be added with the necessary approval. If in doubt please liaise with the application owner."}
        	wl("Create Device AD Groups (_I,_U): created AD group $ADGroupNamePrefix$($InstallCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
        wl("Create Device AD Groups (_I,_U): New-ADGroup ""$ADGroupNamePrefix$($UninstallCollectionName)"" -Path ""$ADOUPath"" -GroupCategory Security -GroupScope Global -Description ""$localizedName""")
		try{
			New-ADGroup "$ADGroupNamePrefix$($UninstallCollectionName)" -Path "$ADOUPath" -GroupCategory Security -GroupScope Global -Description "$localizedName" 
        	wl("Create Device AD Groups (_I,_U): created AD group $ADGroupNamePrefix$($UninstallCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Create User AD Group (_A)
    if($cbCreateUserADGroup.checked -And $returnerror -eq ""){
        wl("Create User AD Group (_A): New-ADGroup ""$ADGroupNamePrefix$($UserCollectionName)"" -Path ""$ADOUPath"" -GroupCategory Security -GroupScope Global -Description ""$localizedName""")
		try{
			New-ADGroup "$ADGroupNamePrefix$($UserCollectionName)" -Path "$ADOUPath" -GroupCategory Security -GroupScope Global -Description "$localizedName"`
-OtherAttributes @{info="This group makes available $($tbUnikeyRef.text) `"$PackageName`" for users specified - Please do not add computer to this group. Users should only be added with the necessary approval. If in doubt please liaise with the application owner."}
        	wl("Create User AD Group (_A): created AD group $ADGroupNamePrefix$($UserCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Add AD Groups Queries to Device Collections
    if($cbAddADGroupsQueriesToDeviceCollections.checked -And $returnerror -eq ""){
        wl("Add AD Groups Queries to Device Collections: Add-CMDeviceCollectionQueryMembershipRule -CollectionName ""$InstallCollectionName"" -RuleName ""$ADGroupNamePrefix$($InstallCollectionName)"" -QueryExpression ""$ADGroupQuery$ADGroupNamePrefix$($InstallCollectionName)'""")
		try{
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$InstallCollectionName" -RuleName "$ADGroupNamePrefix$($InstallCollectionName)" -QueryExpression "$ADGroupQuery$ADGroupNamePrefix$($InstallCollectionName)'"
        	wl("Add AD Groups Queries to Device Collections: created query for AD group $ADGroupNamePrefix$($InstallCollectionName) in $InstallCollectionName")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
        wl("Add AD Groups Queries to Device Collections: Add-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UninstallCollectionName"" -RuleName ""$ADGroupNamePrefix$($UninstallCollectionName)"" -QueryExpression ""$ADGroupQuery$ADGroupNamePrefix$($UninstallCollectionName)'""")
		try{
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$UninstallCollectionName" -RuleName "$ADGroupNamePrefix$($UninstallCollectionName)" -QueryExpression "$ADGroupQuery$ADGroupNamePrefix$($UninstallCollectionName)'"
        	wl("Add AD Groups Queries to Device Collections: created query for AD group $ADGroupNamePrefix$($UninstallCollectionName) in $UninstallCollectionName")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Add AD Group Query To User Collection
    if($cbAddADGroupQueryToUserCollection.checked -And $returnerror -eq ""){
        wl("Add AD Group Query To User Collection: Add-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UserCollectionName"" -RuleName ""$ADGroupNamePrefix$($UserCollectionName)"" -QueryExpression ""$ADGroupUserQuery$ADGroupNamePrefix$($UserCollectionName)'""")
		try{
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$UserCollectionName" -RuleName "$ADGroupNamePrefix$($UserCollectionName)" -QueryExpression "$ADGroupUserQuery$ADGroupNamePrefix$($UserCollectionName)'"
        	wl("Add AD Group Query To User Collection: created query for AD group $ADGroupNamePrefix$($UserCollectionName) in $UserCollectionName")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Add Test Machines To Device Collection
    if($cbAddTestMachineToDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("Add Test Machines To Device Collection: Add-CMDeviceCollectionDirectMembershipRule -CollectionName ""$InstallCollectionName"" -Resource `$(Get-CMDevice -Name $tm)")
    		try{
    			Add-CMDeviceCollectionDirectMembershipRule -CollectionName "$InstallCollectionName" -Resource $(Get-CMDevice -Name $tm)
        	    wl("Add Test Machines To Device Collection: added $tm to $InstallCollectionName")
    		}catch{
    			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
        }
    }
    #Remove Test Machines From Device Collection
    if($cbRemoveTestMachineFromDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("Remove Test Machines From Device Collection: Remove-CMDeviceCollectionDirectMembershipRule -CollectionName ""$InstallCollectionName"" -ResourceName $tm -Force")
    		try{
    			Remove-CMDeviceCollectionDirectMembershipRule -CollectionName "$InstallCollectionName" -ResourceName $tm -Force
        	    wl("Remove Test Machines From Device Collection: removed $tm from $InstallCollectionName")
    		}catch{
    			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
        }
    }
    #Add Test Users To User Collection
    if($cbAddTestUserToUserCollection.checked -And $returnerror -eq ""){
        foreach($tu in $testuser){
            wl("Add Test Users To User Collection: Add-CMUserCollectionDirectMembershipRule -CollectionName ""$UserCollectionName"" -Resource `$(Get-CMUser -Name ""$DomainPrefix\$tu"")")
    		try{
    			Add-CMUserCollectionDirectMembershipRule -CollectionName "$UserCollectionName" -Resource $(Get-CMUser -Name "$DomainPrefix\$tu")
        	    wl("Add Test Users To User Collection: added $tu to $InstallCollectionName")
    		}catch{
    			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
        }
    }
    #Remove Test Users From User Collection
    if($cbRemoveTestUserFromUserCollection.checked -And $returnerror -eq ""){
        foreach($tu in $testuser){
            wl("Remove Test Users From User Collection: Remove-CMUserCollectionDirectMembershipRule -CollectionName ""$UserCollectionName"" -Resource `$(Get-CMUser -Name ""$DomainPrefix\$tu"") -Force")
    		try{
    			Remove-CMUserCollectionDirectMembershipRule -CollectionName "$UserCollectionName" -Resource $(Get-CMUser -Name "$DomainPrefix\$tu") -Force
        	    wl("Remove Test Users From User Collection: removed $tu from $UserCollectionName")
    		}catch{
    			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
        }
    }
    #Update Content On Distribution Points
    if($cbUpdateContent.checked -And $returnerror -eq ""){
        wl("Update Content On Distribution Points: Enumirating Deployment Types for $PackageName application")
		try{
            foreach($dt in $(Get-CMDeploymentType -ApplicationName "$PackageName")){
               wl("Update Content On Distribution Points: Update-CMDistributionPoint -ApplicationName `"$PackageName`" -DeploymentTypeName `"$($dt.LocalizedDisplayName)`"")
               Update-CMDistributionPoint -ApplicationName "$PackageName" -DeploymentTypeName "$($dt.LocalizedDisplayName)"
        	   wl("Update Content On Distribution Points: updated distributed content for $($dt.LocalizedDisplayName)")
            }            
			
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    if($returnerror -eq ""){
        PresetNotSet
        wl("Completed Executing Actions")
    }
}


#### Main


ClearForm
PresetNotSet


[void]$Form.ShowDialog()
