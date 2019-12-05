
#### Defaults


$SiteCode = "SiteCode" 
$SiteServer = "SiteServer" 
$PackageRepository = "UNC path to a folder with packages from where they are published"
$NewPackageLocation = "Object Path in SCCM to Applications"
$DistributionPointGroups = "Distribution Groups"
$InstallDeviceCollectionLocation = "Object Path to Install Device Collection"
$UninstallDeviceCollectionLocation = "Object Path to Uninstall Device Collection"
$LimitingDeviceCollectionName = "e.g. All Windows 10"
$LimitingUserCollectionName = "e.g. All Users"
$UserCollectionLocation = "Object Path to User Collection"
$ADOUPath = "AD Path to Security Groups"
$DomainPrefix = "DomainPrefix"  # Used for AD query below
$ADGroupQuery = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SystemGroupName = '$($DomainPrefix)\\"
$ADGroupUserQuery = "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,SMS_R_USER.Name,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_User where SMS_R_User.UserGroupName = '$($DomainPrefix)\\"
$DetectionKeyLocation = "SOFTWARE\CompanyName\Installed" # Standard Audit key is located HKLM\SOFTWARE\CompanyName\Installed\$PackageName, key name Installed, value Date and time of installation
$ADGroupNamePrefix = "e.g. SCCM_"
$ApplicationFoldersSelectedItem = "Preloaded Object Path to Applications"
$PreloadApplicationFolders = "0"
$OnlyPreselectFolder = "0"
$LoadApplications = "0"

#### Defaults

function SetDefaultVariables{
    $tbTestUser.text = "$env:UserName"
    $tbTestMachine.text = "$env:COMPUTERNAME"
    $script:SiteCode = "SiteCode" # Site code 
    $script:SiteServer = "SiteServer" # SMS Provider machine name
    $script:PackageRepository = "\\SCCMSERVER"
    $script:NewPackageLocation = ""
    $script:DistributionPointGroups = "All DP's"
    $script:InstallDeviceCollectionLocation = ""
    $script:UninstallDeviceCollectionLocation = ""
    $script:LimitingDeviceCollectionName = "All Windows 10 Machines"
    $script:LimitingUserCollectionName = "All Users (Win10)"
    $script:UserCollectionLocation = ""
    $script:ADOUPath = "OU=Application Deployment,DC=company,DC=com"
    $script:DomainPrefix = "DOMAIN"  # Used for AD query below
    $script:ADGroupQuery = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SystemGroupName = '$($DomainPrefix)\\"
    $script:ADGroupUserQuery = "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,SMS_R_USER.Name,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_User where SMS_R_User.UserGroupName = '$($DomainPrefix)\\"
    $script:DetectionKeyLocation = "SOFTWARE\CompanyName\Installed" # Standard Audit key is located HKLM\SOFTWARE\CompanyName\Installed\$PackageName, key name Installed, value Date and time of installation
    $script:ADGroupNamePrefix = "SCCM_"
    $script:ApplicationFoldersSelectedItem = ""
    $script:PreloadApplicationFolders = "0"
    $script:OnlyPreselectFolder = "1"
    $script:LoadApplications = "1"
}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1150,860'
$Form.text                       = "Publish Package to SCCM"
$Form.TopMost                    = $false

$LabelMainForm1                  = New-Object system.Windows.Forms.Label
$LabelMainForm1.text             = "PackageName"
$LabelMainForm1.AutoSize         = $true
$LabelMainForm1.width            = 254
$LabelMainForm1.height           = 10
$LabelMainForm1.location         = New-Object System.Drawing.Point(19,24)
$LabelMainForm1.Font             = 'Microsoft Sans Serif,10'

$LabelMainForm2                  = New-Object system.Windows.Forms.Label
$LabelMainForm2.text             = "Publisher"
$LabelMainForm2.AutoSize         = $true
$LabelMainForm2.width            = 254
$LabelMainForm2.height           = 10
$LabelMainForm2.location         = New-Object System.Drawing.Point(20,50)
$LabelMainForm2.Font             = 'Microsoft Sans Serif,10'

$LabelMainForm3                  = New-Object system.Windows.Forms.Label
$LabelMainForm3.text             = "ApplicationName"
$LabelMainForm3.AutoSize         = $true
$LabelMainForm3.width            = 254
$LabelMainForm3.height           = 10
$LabelMainForm3.location         = New-Object System.Drawing.Point(20,80)
$LabelMainForm3.Font             = 'Microsoft Sans Serif,10'

$LabelMainForm4                  = New-Object system.Windows.Forms.Label
$LabelMainForm4.text             = "Version"
$LabelMainForm4.AutoSize         = $true
$LabelMainForm4.width            = 254
$LabelMainForm4.height           = 10
$LabelMainForm4.location         = New-Object System.Drawing.Point(20,110)
$LabelMainForm4.Font             = 'Microsoft Sans Serif,10'

$gbMainForm1                     = New-Object system.Windows.Forms.Groupbox
$gbMainForm1.height              = 45
$gbMainForm1.width               = 540
$gbMainForm1.text                = "Deployment Type"
$gbMainForm1.location            = New-Object System.Drawing.Point(600,710)

$tbPackageName                   = New-Object system.Windows.Forms.TextBox
$tbPackageName.multiline         = $false
$tbPackageName.width             = 400
$tbPackageName.height            = 20
$tbPackageName.Anchor            = 'top,right,left'
$tbPackageName.location          = New-Object System.Drawing.Point(135,20)
$tbPackageName.Font              = 'Microsoft Sans Serif,10'

$tbPublisher                     = New-Object system.Windows.Forms.TextBox
$tbPublisher.multiline           = $false
$tbPublisher.width               = 400
$tbPublisher.height              = 20
$tbPublisher.Anchor              = 'top,right,left'
$tbPublisher.location            = New-Object System.Drawing.Point(135,50)
$tbPublisher.Font                = 'Microsoft Sans Serif,10'

$tbApplicationName               = New-Object system.Windows.Forms.TextBox
$tbApplicationName.multiline     = $false
$tbApplicationName.width         = 400
$tbApplicationName.height        = 20
$tbApplicationName.Anchor        = 'top,right,left'
$tbApplicationName.location      = New-Object System.Drawing.Point(135,80)
$tbApplicationName.Font          = 'Microsoft Sans Serif,10'

$tbVersion                       = New-Object system.Windows.Forms.TextBox
$tbVersion.multiline             = $false
$tbVersion.width                 = 400
$tbVersion.height                = 20
$tbVersion.Anchor                = 'top,right,left'
$tbVersion.location              = New-Object System.Drawing.Point(135,110)
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

$btnStart                        = New-Object system.Windows.Forms.Button
$btnStart.text                   = "Start Executing Actions"
$btnStart.width                  = 180
$btnStart.height                 = 30
$btnStart.location               = New-Object System.Drawing.Point(900,300)
$btnStart.Font                   = 'Microsoft Sans Serif,10'

$errPackageName                  = New-Object system.Windows.Forms.Label
$errPackageName.text             = "*"
$errPackageName.AutoSize         = $true
$errPackageName.visible          = $false
$errPackageName.width            = 25
$errPackageName.height           = 10
$errPackageName.location         = New-Object System.Drawing.Point(125,20)
$errPackageName.Font             = 'Microsoft Sans Serif,10'
$errPackageName.ForeColor        = "#ff0000"

$errPublisher                    = New-Object system.Windows.Forms.Label
$errPublisher.text               = "*"
$errPublisher.AutoSize           = $true
$errPublisher.visible            = $false
$errPublisher.width              = 25
$errPublisher.height             = 10
$errPublisher.location           = New-Object System.Drawing.Point(125,50)
$errPublisher.Font               = 'Microsoft Sans Serif,10'
$errPublisher.ForeColor          = "#ff0000"

$errApplicationName              = New-Object system.Windows.Forms.Label
$errApplicationName.text         = "*"
$errApplicationName.AutoSize     = $true
$errApplicationName.visible      = $false
$errApplicationName.width        = 25
$errApplicationName.height       = 10
$errApplicationName.location     = New-Object System.Drawing.Point(125,80)
$errApplicationName.Font         = 'Microsoft Sans Serif,10'
$errApplicationName.ForeColor    = "#ff0000"

$errVersion                      = New-Object system.Windows.Forms.Label
$errVersion.text                 = "*"
$errVersion.AutoSize             = $true
$errVersion.visible              = $false
$errVersion.width                = 25
$errVersion.height               = 10
$errVersion.Anchor               = 'top,right'
$errVersion.location             = New-Object System.Drawing.Point(125,110)
$errVersion.Font                 = 'Microsoft Sans Serif,10'
$errVersion.ForeColor            = "#ff0000"

$LabelMainForm6                  = New-Object system.Windows.Forms.Label
$LabelMainForm6.text             = "Test Machines (;)"
$LabelMainForm6.AutoSize         = $true
$LabelMainForm6.width            = 25
$LabelMainForm6.height           = 10
$LabelMainForm6.location         = New-Object System.Drawing.Point(20,20)
$LabelMainForm6.Font             = 'Microsoft Sans Serif,10'

$tbTestMachine                   = New-Object system.Windows.Forms.TextBox
$tbTestMachine.multiline         = $false
$tbTestMachine.width             = 400
$tbTestMachine.height            = 20
$tbTestMachine.Anchor            = 'top,right,left'
$tbTestMachine.location          = New-Object System.Drawing.Point(135,20)
$tbTestMachine.Font              = 'Microsoft Sans Serif,10'

$LabelMainForm7                  = New-Object system.Windows.Forms.Label
$LabelMainForm7.text             = "Test Users (;)"
$LabelMainForm7.AutoSize         = $true
$LabelMainForm7.width            = 25
$LabelMainForm7.height           = 10
$LabelMainForm7.location         = New-Object System.Drawing.Point(20,50)
$LabelMainForm7.Font             = 'Microsoft Sans Serif,10'

$tbTestUser                      = New-Object system.Windows.Forms.TextBox
$tbTestUser.multiline            = $false
$tbTestUser.width                = 400
$tbTestUser.height               = 20
$tbTestUser.Anchor               = 'top,right,left'
$tbTestUser.location             = New-Object System.Drawing.Point(135,50)
$tbTestUser.Font                 = 'Microsoft Sans Serif,10'

$lbLog                           = New-Object system.Windows.Forms.ListBox
$lbLog.text                      = "Log"
$lbLog.width                     = 560
$lbLog.height                    = 170
$lbLog.Anchor                    = 'top,right,bottom,left'
$lbLog.location                  = New-Object System.Drawing.Point(10,10)

$ddApplications                  = New-Object system.Windows.Forms.ComboBox
$ddApplications.text             = "Click Get Applications for Selected Folder button to Populate this list"
$ddApplications.width            = 560
$ddApplications.height           = 20
$ddApplications.visible          = $true
$ddApplications.enabled          = $false
$ddApplications.Anchor           = 'right,bottom,left'
$ddApplications.location         = New-Object System.Drawing.Point(10,210)
$ddApplications.Font             = 'Microsoft Sans Serif,10'

$btnRefreshApplications          = New-Object system.Windows.Forms.Button
$btnRefreshApplications.text     = "Get Applications for Selected Folder"
$btnRefreshApplications.width    = 270
$btnRefreshApplications.height   = 30
$btnRefreshApplications.visible  = $true
$btnRefreshApplications.enabled  = $false
$btnRefreshApplications.Anchor   = 'bottom,left'
$btnRefreshApplications.location  = New-Object System.Drawing.Point(300,240)
$btnRefreshApplications.Font     = 'Microsoft Sans Serif,10'

$btnLoadApplication              = New-Object system.Windows.Forms.Button
$btnLoadApplication.text         = "Load Application"
$btnLoadApplication.width        = 150
$btnLoadApplication.height       = 30
$btnLoadApplication.visible      = $true
$btnLoadApplication.enabled      = $false
$btnLoadApplication.Anchor       = 'bottom,left'
$btnLoadApplication.location     = New-Object System.Drawing.Point(10,280)
$btnLoadApplication.Font         = 'Microsoft Sans Serif,10'

$ddApplicationFolders            = New-Object system.Windows.Forms.ComboBox
$ddApplicationFolders.text       = "Click Refresh Application Folders button to Populate this list"
$ddApplicationFolders.width      = 560
$ddApplicationFolders.height     = 20
$ddApplicationFolders.Anchor     = 'right,bottom,left'
$ddApplicationFolders.location   = New-Object System.Drawing.Point(10,180)
$ddApplicationFolders.Font       = 'Microsoft Sans Serif,10'

$btnRefreshApplicationFolders    = New-Object system.Windows.Forms.Button
$btnRefreshApplicationFolders.text  = "Refresh Application Folders"
$btnRefreshApplicationFolders.width  = 200
$btnRefreshApplicationFolders.height  = 30
$btnRefreshApplicationFolders.Anchor  = 'bottom,left'
$btnRefreshApplicationFolders.location  = New-Object System.Drawing.Point(10,240)
$btnRefreshApplicationFolders.Font  = 'Microsoft Sans Serif,10'

$btnRemoveApplication            = New-Object system.Windows.Forms.Button
$btnRemoveApplication.text       = "Remove Application and its Deployments"
$btnRemoveApplication.width      = 310
$btnRemoveApplication.height     = 30
$btnRemoveApplication.visible    = $true
$btnRemoveApplication.enabled    = $false
$btnRemoveApplication.Anchor     = 'bottom,left'
$btnRemoveApplication.location   = New-Object System.Drawing.Point(260,280)
$btnRemoveApplication.Font       = 'Microsoft Sans Serif,10'

$tbUnikeyRef                     = New-Object system.Windows.Forms.TextBox
$tbUnikeyRef.multiline           = $false
$tbUnikeyRef.text                = "PFA-"
$tbUnikeyRef.width               = 400
$tbUnikeyRef.height              = 20
$tbUnikeyRef.Anchor              = 'top,right,left'
$tbUnikeyRef.location            = New-Object System.Drawing.Point(135,140)
$tbUnikeyRef.Font                = 'Microsoft Sans Serif,10'

$lblUnikeyRef                    = New-Object system.Windows.Forms.Label
$lblUnikeyRef.text               = "Unikey Ref"
$lblUnikeyRef.AutoSize           = $true
$lblUnikeyRef.width              = 25
$lblUnikeyRef.height             = 10
$lblUnikeyRef.location           = New-Object System.Drawing.Point(20,140)
$lblUnikeyRef.Font               = 'Microsoft Sans Serif,10'

$GroupboxMainForm2               = New-Object system.Windows.Forms.Groupbox
$GroupboxMainForm2.height        = 170
$GroupboxMainForm2.width         = 540
$GroupboxMainForm2.text          = "Package Properties"
$GroupboxMainForm2.location      = New-Object System.Drawing.Point(600,530)

$GroupboxMainForm3               = New-Object system.Windows.Forms.Groupbox
$GroupboxMainForm3.height        = 80
$GroupboxMainForm3.width         = 540
$GroupboxMainForm3.text          = "Test"
$GroupboxMainForm3.location      = New-Object System.Drawing.Point(600,770)

$GroupboxMainForm4               = New-Object system.Windows.Forms.Groupbox
$GroupboxMainForm4.height        = 320
$GroupboxMainForm4.width         = 580
$GroupboxMainForm4.location      = New-Object System.Drawing.Point(10,530)

$gbActions                       = New-Object system.Windows.Forms.Groupbox
$gbActions.height                = 515
$gbActions.width                 = 850
$gbActions.location              = New-Object System.Drawing.Point(10,10)

$gbCreatePackage                 = New-Object system.Windows.Forms.Groupbox
$gbCreatePackage.height          = 145
$gbCreatePackage.width           = 179
$gbCreatePackage.text            = "CreatePackage"
$gbCreatePackage.location        = New-Object System.Drawing.Point(10,10)

$cbCreatePackage                 = New-Object system.Windows.Forms.CheckBox
$cbCreatePackage.text            = "Create Package"
$cbCreatePackage.AutoSize        = $true
$cbCreatePackage.width           = 95
$cbCreatePackage.height          = 20
$cbCreatePackage.location        = New-Object System.Drawing.Point(10,20)
$cbCreatePackage.Font            = 'Microsoft Sans Serif,10'

$cbMovePackage                   = New-Object system.Windows.Forms.CheckBox
$cbMovePackage.text              = "Move Package"
$cbMovePackage.AutoSize          = $true
$cbMovePackage.width             = 95
$cbMovePackage.height            = 20
$cbMovePackage.location          = New-Object System.Drawing.Point(10,40)
$cbMovePackage.Font              = 'Microsoft Sans Serif,10'

$cbCreateDeploymentType          = New-Object system.Windows.Forms.CheckBox
$cbCreateDeploymentType.text     = "Create Deployment Type"
$cbCreateDeploymentType.AutoSize  = $true
$cbCreateDeploymentType.width    = 95
$cbCreateDeploymentType.height   = 20
$cbCreateDeploymentType.location  = New-Object System.Drawing.Point(10,60)
$cbCreateDeploymentType.Font     = 'Microsoft Sans Serif,10'

$cbDistributeContent             = New-Object system.Windows.Forms.CheckBox
$cbDistributeContent.text        = "Distribute Content"
$cbDistributeContent.AutoSize    = $true
$cbDistributeContent.width       = 95
$cbDistributeContent.height      = 20
$cbDistributeContent.location    = New-Object System.Drawing.Point(10,80)
$cbDistributeContent.Font        = 'Microsoft Sans Serif,10'

$gbCreateCollections             = New-Object system.Windows.Forms.Groupbox
$gbCreateCollections.height      = 145
$gbCreateCollections.width       = 226
$gbCreateCollections.text        = "Create Collections"
$gbCreateCollections.location    = New-Object System.Drawing.Point(200,10)

$cbCreateInstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbCreateInstallDeviceCollection.text  = "Create Install Device Collection"
$cbCreateInstallDeviceCollection.AutoSize  = $true
$cbCreateInstallDeviceCollection.width  = 95
$cbCreateInstallDeviceCollection.height  = 20
$cbCreateInstallDeviceCollection.location  = New-Object System.Drawing.Point(10,20)
$cbCreateInstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbCreateUninstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbCreateUninstallDeviceCollection.text  = "Create Uninstall Device Collection"
$cbCreateUninstallDeviceCollection.AutoSize  = $true
$cbCreateUninstallDeviceCollection.width  = 95
$cbCreateUninstallDeviceCollection.height  = 20
$cbCreateUninstallDeviceCollection.location  = New-Object System.Drawing.Point(10,60)
$cbCreateUninstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbCreateAvailableUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbCreateAvailableUserCollection.text  = "Create Available User Collection"
$cbCreateAvailableUserCollection.AutoSize  = $true
$cbCreateAvailableUserCollection.width  = 95
$cbCreateAvailableUserCollection.height  = 20
$cbCreateAvailableUserCollection.location  = New-Object System.Drawing.Point(10,100)
$cbCreateAvailableUserCollection.Font  = 'Microsoft Sans Serif,10'

$gbDeployToCollections           = New-Object system.Windows.Forms.Groupbox
$gbDeployToCollections.height    = 145
$gbDeployToCollections.width     = 400
$gbDeployToCollections.text      = "Deploy To Collections (Create Deployment)"
$gbDeployToCollections.location  = New-Object System.Drawing.Point(440,10)

$cbDeployToInstallDeviceCollectionHiddenRequired   = New-Object system.Windows.Forms.CheckBox
$cbDeployToInstallDeviceCollectionHiddenRequired.text  = "Deploy To Install Device Collection (Hidden, Required)"
$cbDeployToInstallDeviceCollectionHiddenRequired.AutoSize  = $true
$cbDeployToInstallDeviceCollectionHiddenRequired.width  = 95
$cbDeployToInstallDeviceCollectionHiddenRequired.height  = 20
$cbDeployToInstallDeviceCollectionHiddenRequired.location  = New-Object System.Drawing.Point(10,20)
$cbDeployToInstallDeviceCollectionHiddenRequired.Font  = 'Microsoft Sans Serif,10'

$cbDeployToUninstallDeviceCollectionHiddenRequired   = New-Object system.Windows.Forms.CheckBox
$cbDeployToUninstallDeviceCollectionHiddenRequired.text  = "Deploy To Uninstall Device Collection (Hidden, Required)"
$cbDeployToUninstallDeviceCollectionHiddenRequired.AutoSize  = $true
$cbDeployToUninstallDeviceCollectionHiddenRequired.width  = 95
$cbDeployToUninstallDeviceCollectionHiddenRequired.height  = 20
$cbDeployToUninstallDeviceCollectionHiddenRequired.location  = New-Object System.Drawing.Point(10,40)
$cbDeployToUninstallDeviceCollectionHiddenRequired.Font  = 'Microsoft Sans Serif,10'

$cbDeployToUserCollectionAvailable   = New-Object system.Windows.Forms.CheckBox
$cbDeployToUserCollectionAvailable.text  = "Deploy To User Collection Available"
$cbDeployToUserCollectionAvailable.AutoSize  = $true
$cbDeployToUserCollectionAvailable.width  = 95
$cbDeployToUserCollectionAvailable.height  = 20
$cbDeployToUserCollectionAvailable.location  = New-Object System.Drawing.Point(10,60)
$cbDeployToUserCollectionAvailable.Font  = 'Microsoft Sans Serif,10'

$gbCreateADGroups                = New-Object system.Windows.Forms.Groupbox
$gbCreateADGroups.height         = 85
$gbCreateADGroups.width          = 250
$gbCreateADGroups.text           = "Create AD Groups"
$gbCreateADGroups.location       = New-Object System.Drawing.Point(10,160)

$cbCreateInstallADGroup          = New-Object system.Windows.Forms.CheckBox
$cbCreateInstallADGroup.text     = "Create Install AD Group (_I)"
$cbCreateInstallADGroup.AutoSize  = $true
$cbCreateInstallADGroup.width    = 95
$cbCreateInstallADGroup.height   = 20
$cbCreateInstallADGroup.location  = New-Object System.Drawing.Point(10,20)
$cbCreateInstallADGroup.Font     = 'Microsoft Sans Serif,10'

$cbCreateUninstallADGroup        = New-Object system.Windows.Forms.CheckBox
$cbCreateUninstallADGroup.text   = "Create Uninstall AD Group (_U)"
$cbCreateUninstallADGroup.AutoSize  = $true
$cbCreateUninstallADGroup.width  = 95
$cbCreateUninstallADGroup.height  = 20
$cbCreateUninstallADGroup.location  = New-Object System.Drawing.Point(10,40)
$cbCreateUninstallADGroup.Font   = 'Microsoft Sans Serif,10'

$cbCreateUserAvailableADGroup    = New-Object system.Windows.Forms.CheckBox
$cbCreateUserAvailableADGroup.text  = "Create User Available AD Group (_A)"
$cbCreateUserAvailableADGroup.AutoSize  = $true
$cbCreateUserAvailableADGroup.width  = 95
$cbCreateUserAvailableADGroup.height  = 20
$cbCreateUserAvailableADGroup.location  = New-Object System.Drawing.Point(10,60)
$cbCreateUserAvailableADGroup.Font  = 'Microsoft Sans Serif,10'

$gbCreateADGroupsQueries         = New-Object system.Windows.Forms.Groupbox
$gbCreateADGroupsQueries.height  = 85
$gbCreateADGroupsQueries.width   = 340
$gbCreateADGroupsQueries.text    = "Create AD Groups Queries"
$gbCreateADGroupsQueries.location  = New-Object System.Drawing.Point(270,160)

$cbAddADGroupQueryToInstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddADGroupQueryToInstallDeviceCollection.text  = "Add AD Group Query To Install Device Collection"
$cbAddADGroupQueryToInstallDeviceCollection.AutoSize  = $true
$cbAddADGroupQueryToInstallDeviceCollection.width  = 95
$cbAddADGroupQueryToInstallDeviceCollection.height  = 20
$cbAddADGroupQueryToInstallDeviceCollection.location  = New-Object System.Drawing.Point(10,20)
$cbAddADGroupQueryToInstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbAddADGroupQueryToUninstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddADGroupQueryToUninstallDeviceCollection.text  = "Add AD Group Query To Uninstall Device Collection"
$cbAddADGroupQueryToUninstallDeviceCollection.AutoSize  = $true
$cbAddADGroupQueryToUninstallDeviceCollection.width  = 95
$cbAddADGroupQueryToUninstallDeviceCollection.height  = 20
$cbAddADGroupQueryToUninstallDeviceCollection.location  = New-Object System.Drawing.Point(10,40)
$cbAddADGroupQueryToUninstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbAddADGroupQueryToAvailableUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddADGroupQueryToAvailableUserCollection.text  = "Add AD Group Query To Available User Collection"
$cbAddADGroupQueryToAvailableUserCollection.AutoSize  = $true
$cbAddADGroupQueryToAvailableUserCollection.width  = 95
$cbAddADGroupQueryToAvailableUserCollection.height  = 20
$cbAddADGroupQueryToAvailableUserCollection.location  = New-Object System.Drawing.Point(10,60)
$cbAddADGroupQueryToAvailableUserCollection.Font  = 'Microsoft Sans Serif,10'

$gbExtra                         = New-Object system.Windows.Forms.Groupbox
$gbExtra.height                  = 85
$gbExtra.width                   = 220
$gbExtra.text                    = "Extra"
$gbExtra.location                = New-Object System.Drawing.Point(620,160)

$cbUpdateContentOnDPs            = New-Object system.Windows.Forms.CheckBox
$cbUpdateContentOnDPs.text       = "Update Content On DPs"
$cbUpdateContentOnDPs.AutoSize   = $true
$cbUpdateContentOnDPs.width      = 95
$cbUpdateContentOnDPs.height     = 20
$cbUpdateContentOnDPs.location   = New-Object System.Drawing.Point(10,20)
$cbUpdateContentOnDPs.Font       = 'Microsoft Sans Serif,10'

$gbTestMachines                  = New-Object system.Windows.Forms.Groupbox
$gbTestMachines.height           = 105
$gbTestMachines.width            = 410
$gbTestMachines.text             = "Test Machines"
$gbTestMachines.location         = New-Object System.Drawing.Point(10,250)

$cbAddTestMachinesToInstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddTestMachinesToInstallDeviceCollection.text  = "Add Test Machines To Install Device Collection"
$cbAddTestMachinesToInstallDeviceCollection.AutoSize  = $true
$cbAddTestMachinesToInstallDeviceCollection.width  = 95
$cbAddTestMachinesToInstallDeviceCollection.height  = 20
$cbAddTestMachinesToInstallDeviceCollection.location  = New-Object System.Drawing.Point(10,20)
$cbAddTestMachinesToInstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveTestMachinesFromInstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveTestMachinesFromInstallDeviceCollection.text  = "Remove Test Machines From Install Device Collection"
$cbRemoveTestMachinesFromInstallDeviceCollection.AutoSize  = $true
$cbRemoveTestMachinesFromInstallDeviceCollection.width  = 95
$cbRemoveTestMachinesFromInstallDeviceCollection.height  = 20
$cbRemoveTestMachinesFromInstallDeviceCollection.location  = New-Object System.Drawing.Point(10,40)
$cbRemoveTestMachinesFromInstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbAddTestMachinesToUninstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddTestMachinesToUninstallDeviceCollection.text  = "Add Test Machines To Uninstall Device Collection"
$cbAddTestMachinesToUninstallDeviceCollection.AutoSize  = $true
$cbAddTestMachinesToUninstallDeviceCollection.width  = 95
$cbAddTestMachinesToUninstallDeviceCollection.height  = 20
$cbAddTestMachinesToUninstallDeviceCollection.location  = New-Object System.Drawing.Point(10,60)
$cbAddTestMachinesToUninstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveTestMachinesFromUninstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveTestMachinesFromUninstallDeviceCollection.text  = "Remove Test Machines From Uninstall Device Collection"
$cbRemoveTestMachinesFromUninstallDeviceCollection.AutoSize  = $true
$cbRemoveTestMachinesFromUninstallDeviceCollection.width  = 95
$cbRemoveTestMachinesFromUninstallDeviceCollection.height  = 20
$cbRemoveTestMachinesFromUninstallDeviceCollection.location  = New-Object System.Drawing.Point(10,80)
$cbRemoveTestMachinesFromUninstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$gbTestUsers                     = New-Object system.Windows.Forms.Groupbox
$gbTestUsers.height              = 105
$gbTestUsers.width               = 400
$gbTestUsers.text                = "Test Users"
$gbTestUsers.location            = New-Object System.Drawing.Point(440,250)

$cbAddTestUsersToAvailableUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbAddTestUsersToAvailableUserCollection.text  = "Add Test Users To Available User Collection"
$cbAddTestUsersToAvailableUserCollection.AutoSize  = $true
$cbAddTestUsersToAvailableUserCollection.width  = 95
$cbAddTestUsersToAvailableUserCollection.height  = 20
$cbAddTestUsersToAvailableUserCollection.location  = New-Object System.Drawing.Point(10,20)
$cbAddTestUsersToAvailableUserCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveTestUsersFromAvailableUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveTestUsersFromAvailableUserCollection.text  = "Remove Test Users From Available User Collection"
$cbRemoveTestUsersFromAvailableUserCollection.AutoSize  = $true
$cbRemoveTestUsersFromAvailableUserCollection.width  = 95
$cbRemoveTestUsersFromAvailableUserCollection.height  = 20
$cbRemoveTestUsersFromAvailableUserCollection.location  = New-Object System.Drawing.Point(10,40)
$cbRemoveTestUsersFromAvailableUserCollection.Font  = 'Microsoft Sans Serif,10'

$gbRemoveCollections             = New-Object system.Windows.Forms.Groupbox
$gbRemoveCollections.height      = 145
$gbRemoveCollections.width       = 420
$gbRemoveCollections.text        = "Remove Collections"
$gbRemoveCollections.location    = New-Object System.Drawing.Point(10,360)

$cbRemoveInstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveInstallDeviceCollection.text  = "Remove Install Device Collection"
$cbRemoveInstallDeviceCollection.AutoSize  = $true
$cbRemoveInstallDeviceCollection.width  = 95
$cbRemoveInstallDeviceCollection.height  = 20
$cbRemoveInstallDeviceCollection.location  = New-Object System.Drawing.Point(10,20)
$cbRemoveInstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveUninstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveUninstallDeviceCollection.text  = "Remove Uninstall Device Collection"
$cbRemoveUninstallDeviceCollection.AutoSize  = $true
$cbRemoveUninstallDeviceCollection.width  = 95
$cbRemoveUninstallDeviceCollection.height  = 20
$cbRemoveUninstallDeviceCollection.location  = New-Object System.Drawing.Point(10,40)
$cbRemoveUninstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveDeploymentToInstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveDeploymentToInstallDeviceCollection.text  = "Remove Deployment To Install Device Collection"
$cbRemoveDeploymentToInstallDeviceCollection.AutoSize  = $true
$cbRemoveDeploymentToInstallDeviceCollection.width  = 95
$cbRemoveDeploymentToInstallDeviceCollection.height  = 20
$cbRemoveDeploymentToInstallDeviceCollection.location  = New-Object System.Drawing.Point(10,80)
$cbRemoveDeploymentToInstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveDeploymentToUninstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveDeploymentToUninstallDeviceCollection.text  = "Remove Deployment To Uninstall Device Collection"
$cbRemoveDeploymentToUninstallDeviceCollection.AutoSize  = $true
$cbRemoveDeploymentToUninstallDeviceCollection.width  = 95
$cbRemoveDeploymentToUninstallDeviceCollection.height  = 20
$cbRemoveDeploymentToUninstallDeviceCollection.location  = New-Object System.Drawing.Point(10,100)
$cbRemoveDeploymentToUninstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveDeploymentToUserCollectionAvailable   = New-Object system.Windows.Forms.CheckBox
$cbRemoveDeploymentToUserCollectionAvailable.text  = "Remove Deployment To User Collection (Available)"
$cbRemoveDeploymentToUserCollectionAvailable.AutoSize  = $true
$cbRemoveDeploymentToUserCollectionAvailable.width  = 95
$cbRemoveDeploymentToUserCollectionAvailable.height  = 20
$cbRemoveDeploymentToUserCollectionAvailable.location  = New-Object System.Drawing.Point(10,120)
$cbRemoveDeploymentToUserCollectionAvailable.Font  = 'Microsoft Sans Serif,10'

$gbRemoveADGroups                = New-Object system.Windows.Forms.Groupbox
$gbRemoveADGroups.height         = 145
$gbRemoveADGroups.width          = 400
$gbRemoveADGroups.text           = "Remove AD Groups"
$gbRemoveADGroups.location       = New-Object System.Drawing.Point(440,360)

$cbRemoveInstallADGroup          = New-Object system.Windows.Forms.CheckBox
$cbRemoveInstallADGroup.text     = "Remove Install AD Group (_I)"
$cbRemoveInstallADGroup.AutoSize  = $true
$cbRemoveInstallADGroup.width    = 95
$cbRemoveInstallADGroup.height   = 20
$cbRemoveInstallADGroup.location  = New-Object System.Drawing.Point(10,20)
$cbRemoveInstallADGroup.Font     = 'Microsoft Sans Serif,10'

$cbRemoveUninstallADGroup        = New-Object system.Windows.Forms.CheckBox
$cbRemoveUninstallADGroup.text   = "Remove Uninstall AD Group (_U)"
$cbRemoveUninstallADGroup.AutoSize  = $true
$cbRemoveUninstallADGroup.width  = 95
$cbRemoveUninstallADGroup.height  = 20
$cbRemoveUninstallADGroup.location  = New-Object System.Drawing.Point(10,40)
$cbRemoveUninstallADGroup.Font   = 'Microsoft Sans Serif,10'

$cbRemoveUserAvailableADGroup    = New-Object system.Windows.Forms.CheckBox
$cbRemoveUserAvailableADGroup.text  = "Remove User Available AD Group (_A)"
$cbRemoveUserAvailableADGroup.AutoSize  = $true
$cbRemoveUserAvailableADGroup.width  = 95
$cbRemoveUserAvailableADGroup.height  = 20
$cbRemoveUserAvailableADGroup.location  = New-Object System.Drawing.Point(10,60)
$cbRemoveUserAvailableADGroup.Font  = 'Microsoft Sans Serif,10'

$cbRemoveADGroupQueryToInstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveADGroupQueryToInstallDeviceCollection.text  = "Remove AD Group Query To Install Device Collection"
$cbRemoveADGroupQueryToInstallDeviceCollection.AutoSize  = $true
$cbRemoveADGroupQueryToInstallDeviceCollection.width  = 95
$cbRemoveADGroupQueryToInstallDeviceCollection.height  = 20
$cbRemoveADGroupQueryToInstallDeviceCollection.location  = New-Object System.Drawing.Point(10,80)
$cbRemoveADGroupQueryToInstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveADGroupQueryToUninstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveADGroupQueryToUninstallDeviceCollection.text  = "Remove AD Group Query To Uninstall Device Collection"
$cbRemoveADGroupQueryToUninstallDeviceCollection.AutoSize  = $true
$cbRemoveADGroupQueryToUninstallDeviceCollection.width  = 95
$cbRemoveADGroupQueryToUninstallDeviceCollection.height  = 20
$cbRemoveADGroupQueryToUninstallDeviceCollection.location  = New-Object System.Drawing.Point(10,100)
$cbRemoveADGroupQueryToUninstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveADGroupQueryToAvailableUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveADGroupQueryToAvailableUserCollection.text  = "Remove AD Group Query To Available User Collection"
$cbRemoveADGroupQueryToAvailableUserCollection.AutoSize  = $true
$cbRemoveADGroupQueryToAvailableUserCollection.width  = 95
$cbRemoveADGroupQueryToAvailableUserCollection.height  = 20
$cbRemoveADGroupQueryToAvailableUserCollection.location  = New-Object System.Drawing.Point(10,120)
$cbRemoveADGroupQueryToAvailableUserCollection.Font  = 'Microsoft Sans Serif,10'

$ddPresets                       = New-Object system.Windows.Forms.ComboBox
$ddPresets.text                  = "Not Set"
$ddPresets.width                 = 200
$ddPresets.height                = 20
$ddPresets.location              = New-Object System.Drawing.Point(880,20)
$ddPresets.Font                  = 'Microsoft Sans Serif,10'

$cbMoveInstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbMoveInstallDeviceCollection.text  = "Move Install Device Collection"
$cbMoveInstallDeviceCollection.AutoSize  = $true
$cbMoveInstallDeviceCollection.width  = 95
$cbMoveInstallDeviceCollection.height  = 20
$cbMoveInstallDeviceCollection.location  = New-Object System.Drawing.Point(10,40)
$cbMoveInstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbMoveUninstallDeviceCollection   = New-Object system.Windows.Forms.CheckBox
$cbMoveUninstallDeviceCollection.text  = "Move Uninstall Device Collection"
$cbMoveUninstallDeviceCollection.AutoSize  = $true
$cbMoveUninstallDeviceCollection.width  = 95
$cbMoveUninstallDeviceCollection.height  = 20
$cbMoveUninstallDeviceCollection.location  = New-Object System.Drawing.Point(10,80)
$cbMoveUninstallDeviceCollection.Font  = 'Microsoft Sans Serif,10'

$cbMoveAvailableUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbMoveAvailableUserCollection.text  = "Move Available User Collection"
$cbMoveAvailableUserCollection.AutoSize  = $true
$cbMoveAvailableUserCollection.width  = 95
$cbMoveAvailableUserCollection.height  = 20
$cbMoveAvailableUserCollection.location  = New-Object System.Drawing.Point(10,120)
$cbMoveAvailableUserCollection.Font  = 'Microsoft Sans Serif,10'

$cbRemoveAvailableUserCollection   = New-Object system.Windows.Forms.CheckBox
$cbRemoveAvailableUserCollection.text  = "Remove Available User Collection"
$cbRemoveAvailableUserCollection.AutoSize  = $true
$cbRemoveAvailableUserCollection.width  = 95
$cbRemoveAvailableUserCollection.height  = 20
$cbRemoveAvailableUserCollection.location  = New-Object System.Drawing.Point(10,60)
$cbRemoveAvailableUserCollection.Font  = 'Microsoft Sans Serif,10'

$btnOpenLogFile                  = New-Object system.Windows.Forms.Button
$btnOpenLogFile.text             = "Open Log File"
$btnOpenLogFile.width            = 180
$btnOpenLogFile.height           = 30
$btnOpenLogFile.location         = New-Object System.Drawing.Point(900,340)
$btnOpenLogFile.Font             = 'Microsoft Sans Serif,10'

$cbRequestMachineAssignments     = New-Object system.Windows.Forms.CheckBox
$cbRequestMachineAssignments.text  = "Request Machine Assignments"
$cbRequestMachineAssignments.AutoSize  = $true
$cbRequestMachineAssignments.width  = 95
$cbRequestMachineAssignments.height  = 20
$cbRequestMachineAssignments.location  = New-Object System.Drawing.Point(10,40)
$cbRequestMachineAssignments.Font  = 'Microsoft Sans Serif,10'

$LabelMainForm5                  = New-Object system.Windows.Forms.Label
$LabelMainForm5.text             = "for each test machine"
$LabelMainForm5.AutoSize         = $true
$LabelMainForm5.width            = 25
$LabelMainForm5.height           = 10
$LabelMainForm5.location         = New-Object System.Drawing.Point(20,55)
$LabelMainForm5.Font             = 'Microsoft Sans Serif,10'

$btnSettings                     = New-Object system.Windows.Forms.Button
$btnSettings.text                = "Open Settings"
$btnSettings.width               = 180
$btnSettings.height              = 30
$btnSettings.location            = New-Object System.Drawing.Point(900,380)
$btnSettings.Font                = 'Microsoft Sans Serif,10'

$btnLoadSettings                 = New-Object system.Windows.Forms.Button
$btnLoadSettings.text            = "Apply Settings"
$btnLoadSettings.width           = 180
$btnLoadSettings.height          = 30
$btnLoadSettings.location        = New-Object System.Drawing.Point(900,420)
$btnLoadSettings.Font            = 'Microsoft Sans Serif,10'


$btnValidatePackageState         = New-Object system.Windows.Forms.Button
$btnValidatePackageState.text    = "Validate Package State"
$btnValidatePackageState.width   = 180
$btnValidatePackageState.height  = 30
$btnValidatePackageState.location  = New-Object System.Drawing.Point(900,480)
$btnValidatePackageState.Font    = 'Microsoft Sans Serif,10'

$GroupboxMainForm2.controls.AddRange(@($LabelMainForm1,$LabelMainForm2,$LabelMainForm3,$LabelMainForm4,$tbPackageName,$tbPublisher,$tbApplicationName,$tbVersion,$errPackageName,$errPublisher,$errApplicationName,$errVersion,$tbUnikeyRef,$lblUnikeyRef))
$Form.controls.AddRange(@($gbMainForm1,$btnStart,$GroupboxMainForm2,$GroupboxMainForm3,$GroupboxMainForm4,$gbActions,$ddPresets,$btnOpenLogFile,$btnSettings,$btnLoadSettings,$btnValidatePackageState))
$gbMainForm1.controls.AddRange(@($rbMSI,$rbPSAppDeploy,$rbBAT))
$GroupboxMainForm3.controls.AddRange(@($LabelMainForm6,$tbTestMachine,$LabelMainForm7,$tbTestUser))
$GroupboxMainForm4.controls.AddRange(@($lbLog,$ddApplications,$btnRefreshApplications,$btnLoadApplication,$ddApplicationFolders,$btnRefreshApplicationFolders,$btnRemoveApplication))
$gbActions.controls.AddRange(@($gbCreatePackage,$gbCreateCollections,$gbDeployToCollections,$gbCreateADGroups,$gbCreateADGroupsQueries,$gbExtra,$gbTestMachines,$gbTestUsers,$gbRemoveCollections,$gbRemoveADGroups))
$gbCreatePackage.controls.AddRange(@($cbCreatePackage,$cbMovePackage,$cbCreateDeploymentType,$cbDistributeContent))
$gbCreateCollections.controls.AddRange(@($cbCreateInstallDeviceCollection,$cbCreateUninstallDeviceCollection,$cbCreateAvailableUserCollection,$cbMoveInstallDeviceCollection,$cbMoveUninstallDeviceCollection,$cbMoveAvailableUserCollection))
$gbDeployToCollections.controls.AddRange(@($cbDeployToInstallDeviceCollectionHiddenRequired,$cbDeployToUninstallDeviceCollectionHiddenRequired,$cbDeployToUserCollectionAvailable))
$gbCreateADGroups.controls.AddRange(@($cbCreateInstallADGroup,$cbCreateUninstallADGroup,$cbCreateUserAvailableADGroup))
$gbCreateADGroupsQueries.controls.AddRange(@($cbAddADGroupQueryToInstallDeviceCollection,$cbAddADGroupQueryToUninstallDeviceCollection,$cbAddADGroupQueryToAvailableUserCollection))
$gbExtra.controls.AddRange(@($cbUpdateContentOnDPs,$cbRequestMachineAssignments,$LabelMainForm5))
$gbTestMachines.controls.AddRange(@($cbAddTestMachinesToInstallDeviceCollection,$cbRemoveTestMachinesFromInstallDeviceCollection,$cbAddTestMachinesToUninstallDeviceCollection,$cbRemoveTestMachinesFromUninstallDeviceCollection))
$gbTestUsers.controls.AddRange(@($cbAddTestUsersToAvailableUserCollection,$cbRemoveTestUsersFromAvailableUserCollection))
$gbRemoveCollections.controls.AddRange(@($cbRemoveInstallDeviceCollection,$cbRemoveUninstallDeviceCollection,$cbRemoveDeploymentToInstallDeviceCollection,$cbRemoveDeploymentToUninstallDeviceCollection,$cbRemoveDeploymentToUserCollectionAvailable,$cbRemoveAvailableUserCollection))
$gbRemoveADGroups.controls.AddRange(@($cbRemoveInstallADGroup,$cbRemoveUninstallADGroup,$cbRemoveUserAvailableADGroup,$cbRemoveADGroupQueryToInstallDeviceCollection,$cbRemoveADGroupQueryToUninstallDeviceCollection,$cbRemoveADGroupQueryToAvailableUserCollection))

$btnStart.Add_Click({ btnStartClicked })
$Form.Add_Shown({ FormShown })
$btnRefreshApplications.Add_Click({ RefreshApplicationsList })
$btnLoadApplication.Add_Click({ LoadApplication })
$ddApplicationFolders.Add_SelectedValueChanged({ ApplicationFolderSelected })
$btnRefreshApplicationFolders.Add_Click({ RefreshApplicationFoldersClicked })
$btnRemoveApplication.Add_Click({ RemoveLoadedApplicationClicked })
$ddPresets.Add_SelectedValueChanged({ PresetsChanged })
$btnOpenLogFile.Add_Click({ OpenLogFile })
$btnSettings.Add_Click({ FormSettingsOpen })
$btnLoadSettings.Add_Click({ LoadSettings })
$btnValidatePackageState.Add_Click({ btnValidatePackageStateClicked })





$lbLog.HorizontalScrollbar =$true
$rbMSI.checked = $true
[__comobject]$Shell = New-Object -ComObject 'WScript.Shell' -ErrorAction 'SilentlyContinue'

$LogFileLocation = "$($env:TEMP)\PublishPackageToSCCM_$((get-date).tostring('ddMMyyHHmmss')).log"


#### Functions

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

#Write Log
function wl($message){
    try{
        $tmpmessage = "$(Get-Date -Format 'HH:mm:ss') $message"
        $lbLog.Items.Add($tmpmessage)
        $lbLog.SelectedIndex = $lbLog.Items.Count - 1;
        $Form.refresh()
        $tmpmessage|Out-File -FilePath $LogFileLocation -Append
    }catch{
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message)
    }
    return $message
}


function ShowMessageBoxWithError($errmsg){
    if($errmsg -ne ""){
        $UserResponse = $Shell.Popup("$errmsg `n`n Do you wish to continue?",0,"Error",4116)
        if ($UserResponse -eq 6){
            $errmsg = ""
        }
    }
    return $errmsg
}

function ConnectToSCCMAndAD{
    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ConfigurationManager) -eq $null) {
        try{
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
        }
        catch{
            wl("Failed to import SCCM module. Please make sure that you have SCCM Console installed")
            wl($_.Exception.Message)
        }
    }
    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ActiveDirectory) -eq $null) {
        try{
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch{
            wl("Failed to import ActiveDirectory module. Please make sure that you have RSAT installed")
            wl($_.Exception.Message)
        }
    }
    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        try{
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop
        }
        catch{
            wl($_.Exception.Message)
        }
    }
    # Check Package Repository can be connected
    if(![System.IO.Directory]::Exists($PackageRepository)){
        wl("Can't connect to Package repository $($PackageRepository)")
    }
    # Set the current location to be the site code.
    try{
        Set-Location "$($SiteCode):\" -ErrorAction Stop
    }
    catch{
        wl($_.Exception.Message)
    }
}

function FormShown {
    SetVariables
    if($testMode){TestMode}
    LoadSettings 
}


function FormSettingsOpen {
    $FormSettings.Visible = $true
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
        $tmpObj = Get-CMApplication -Name "$appname" -ErrorAction Stop
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

function GetMSIFileName($path){
    $retval = ""
    try{
        cd C:\
        $retval = (dir $path *.msi)[0].Name
    }
    catch{
        wl($_.Exception.Message)
        $retval = ""
    }
    finally{
        Set-Location "$($SiteCode):\"
    }
    return $retval
}

function CheckFileExists($filepath){
    $retval = $false
    try{
        cd C:\
        $retval = Test-Path $filepath
    }
    catch{
        wl($_.Exception.Message)
        $retval = $false
    }
    finally{
        Set-Location "$($SiteCode):\"
    }
    return $retval
}

function OpenLogFile { 
    .$($LogFileLocation)
}

function SetVariables {
    if([System.IO.File]::Exists("$PSScriptRoot\Settings.xml")){
		cd C:\
        $xmlread = [xml](Get-Content "$PSScriptRoot\Settings.xml")
        $tbTestUser.text = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.DefaultUserName)")
        $tbTestMachine.text = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.DefaultTestMachine)")
        $script:SiteCode = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.SiteCode)")
        $script:SiteServer = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.SiteServer)")
        $script:PackageRepository = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.PackageRepository)")
        $script:NewPackageLocation = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.NewPackageLocation)")
        $script:DistributionPointGroups = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.DistributionPointGroups)")
        $script:InstallDeviceCollectionLocation = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.InstallDeviceCollectionLocation)")
        $script:UninstallDeviceCollectionLocation = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.UninstallDeviceCollectionLocation)")
        $script:LimitingDeviceCollectionName = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.LimitingDeviceCollectionName)")
        $script:LimitingUserCollectionName = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.LimitingUserCollectionName)")
        $script:UserCollectionLocation = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.UserCollectionLocation)")
        $script:ADOUPath = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.ADOUPath)")
        $script:DomainPrefix = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.DomainPrefix)")
        $script:ADGroupQuery = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.ADGroupQuery)")
        $script:ADGroupUserQuery = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.ADGroupUserQuery)")
        $script:DetectionKeyLocation = $ExecutionContext.InvokeCommand.ExpandString("$($xmlread.Settings.General.DetectionKeyLocation)")
        $script:ApplicationFoldersSelectedItem = $xmlread.Settings.ActionsOnScriptExecution.SelectedApplicationFolder
        $script:PreloadApplicationFolders = $xmlread.Settings.ActionsOnScriptExecution.PreloadApplicationFolders
        $script:OnlyPreselectFolder = $xmlread.Settings.ActionsOnScriptExecution.OnlyPreselectFolder
        $script:LoadApplications = $xmlread.Settings.ActionsOnScriptExecution.LoadApplications
    }else{
        SetDefaultVariables
    }
    if($SiteCode -eq "SiteCode"){
        FormSettingsOpen
    }
}

function LoadSettings {
    ConnectToSCCMAndAD
    if(ConvertString01ToTrueFalse($PreloadApplicationFolders)){
        RefreshApplicationFoldersClicked
    }
    if(ConvertString01ToTrueFalse($OnlyPreselectFolder)){
        $ddApplicationFolders.Items.Add($ApplicationFoldersSelectedItem)
        $ddApplicationFolders.SelectedItem = $ApplicationFoldersSelectedItem 
        $btnRefreshApplications.enabled = $true
    }
    if(ConvertString01ToTrueFalse($LoadApplications)){
        $btnRefreshApplications.enabled = $true
        RefreshApplicationsList
    }
}


#### Presets


@('Not Set','Create Package','Pass to Regression Testing','Pass to Live','Add Test User Deployment','Add User Deployment To Live') | ForEach-Object {[void] $ddPresets.Items.Add($_)}

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
    $cbMovePackage.checked = $false
    $cbCreateDeploymentType.checked = $false
    $cbDistributeContent.checked = $false
    $cbCreateInstallDeviceCollection.checked = $false
    $cbMoveInstallDeviceCollection.checked = $false
    $cbCreateUninstallDeviceCollection.checked = $false
    $cbMoveUninstallDeviceCollection.checked = $false
    $cbCreateAvailableUserCollection.checked = $false
    $cbMoveAvailableUserCollection.checked = $false
    $cbDeployToInstallDeviceCollectionHiddenRequired.checked = $false
    $cbDeployToUninstallDeviceCollectionHiddenRequired.checked = $false
    $cbDeployToUserCollectionAvailable.checked = $false
    $cbCreateInstallADGroup.checked = $false
    $cbCreateUninstallADGroup.checked = $false
    $cbCreateUserAvailableADGroup.checked = $false
    $cbAddADGroupQueryToInstallDeviceCollection.checked = $false
    $cbAddADGroupQueryToUninstallDeviceCollection.checked = $false
    $cbAddADGroupQueryToAvailableUserCollection.checked = $false
    $cbUpdateContentOnDPs.checked = $false
    $cbRequestMachineAssignments.checked = $false
    $cbAddTestMachinesToInstallDeviceCollection.checked = $false
    $cbRemoveTestMachinesFromInstallDeviceCollection.checked = $false
    $cbAddTestMachinesToUninstallDeviceCollection.checked = $false
    $cbRemoveTestMachinesFromUninstallDeviceCollection.checked = $false
    $cbAddTestUsersToAvailableUserCollection.checked = $false
    $cbRemoveTestUsersFromAvailableUserCollection.checked = $false
    $cbRemoveInstallDeviceCollection.checked = $false
    $cbRemoveUninstallDeviceCollection.checked = $false
    $cbRemoveAvailableUserCollection.checked = $false
    $cbRemoveDeploymentToInstallDeviceCollection.checked = $false
    $cbRemoveDeploymentToUninstallDeviceCollection.checked = $false
    $cbRemoveDeploymentToUserCollectionAvailable.checked = $false
    $cbRemoveInstallADGroup.checked = $false
    $cbRemoveUninstallADGroup.checked = $false
    $cbRemoveUserAvailableADGroup.checked = $false
    $cbRemoveADGroupQueryToInstallDeviceCollection.checked = $false
    $cbRemoveADGroupQueryToUninstallDeviceCollection.checked = $false
    $cbRemoveADGroupQueryToAvailableUserCollection.checked = $false
}
function PresetCreatePackage { 
    PresetNotSet
    $cbCreatePackage.checked = $true
    $cbMovePackage.checked = $true
    $cbCreateDeploymentType.checked = $true
    $cbDistributeContent.checked = $true
    $cbCreateInstallDeviceCollection.checked = $true
    $cbMoveInstallDeviceCollection.checked = $true
    $cbCreateUninstallDeviceCollection.checked = $true
    $cbMoveUninstallDeviceCollection.checked = $true
    $cbDeployToInstallDeviceCollectionHiddenRequired.checked = $true
    $cbDeployToUninstallDeviceCollectionHiddenRequired.checked = $true
    $cbAddTestMachinesToInstallDeviceCollection.checked = $true
    $cbCreateInstallADGroup.checked = $true
    $cbCreateUninstallADGroup.checked = $true
}
function PresetPasstoRegressionTesting {
    PresetNotSet
    $cbRemoveTestMachinesFromInstallDeviceCollection.checked = $true
    $cbAddTestMachinesToUninstallDeviceCollection.checked = $true
}
function PresetPasstoLive { 
    PresetNotSet
    $cbRemoveTestMachinesFromUninstallDeviceCollection.checked = $true
    $cbAddADGroupQueryToInstallDeviceCollection.checked = $true
    $cbAddADGroupQueryToUninstallDeviceCollection.checked = $true
}
function PresetAddTestUserDeployment { 
    $cbCreateAvailableUserCollection.checked = $true
    $cbMoveAvailableUserCollection.checked = $true
    $cbDeployToUserCollectionAvailable.checked = $true
    $cbAddTestUsersToAvailableUserCollection.checked = $true
}
function PresetAddUserDeploymentToLive {
    $cbCreateUserAvailableADGroup.checked = $true
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
    
    return $retValidateForm
}



#### Remove Application


function RemoveLoadedApplicationClicked {
    #$PackageName = $tbPackageName.text
    #if($PackageName -eq ""){
    #    $PackageName=$ddApplications.SelectedItem
    #}
    
    $PackageName=$ddApplications.SelectedItem
    
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
    				Export-CMApplication -Name "$PackageName" -Path "$($env:Temp)\$($PackageName).zip" -Force -ErrorAction Stop
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
            				Remove-CMCollection -Name "$DeploymentCollection" -Force -ErrorAction Stop
            			}
            			catch{
    				        $returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            			}
                    }else{
                        wl("Remove Application: Leaving collection $DeploymentCollection behind as it has other deployments but removing the deployment")
                        wl("Remove Application: Remove-CMApplicationDeployment -Name ""$PackageName"" -CollectionName ""$DeploymentCollection"" -Force")
                        try{				
            				Remove-CMApplicationDeployment -Name "$PackageName" -CollectionName "$DeploymentCollection" -Force -ErrorAction Stop
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
        				Remove-CMApplication -Name "$PackageName" -Force -ErrorAction Stop
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
    $MSIName = ""
    $testmachine = $($tbTestMachine.text.Trim(";")).Split(";")
    $testuser = $($tbTestUser.text.Trim(";")).Split(";")
    
    #calculated variables
    $InstallCollectionName = "$($PackageName)_I"
    $UninstallCollectionName = "$($PackageName)_U"
    $UserCollectionName = "$($PackageName)_A"
    $defaultPackageLocation = "$($PackageRepository)\$PackageName"
    $description = "$($tbUnikeyRef.text) $Publisher $ApplicationName $Version $defaultPackageLocation"
    $localizedName = "$Publisher $ApplicationName $Version"
    
    $returnerror += ValidateForm
    
    if($returnerror -ne ""){
        wl($returnerror)
        $returnerror = ShowMessageBoxWithError("Validation returned error: $returnerror")
    }
    
    if($returnerror -eq ""){
        wl("Started Executing Actions")
    }
    
    $actionname = "Create Package"
    if($cbCreatePackage.checked -And $returnerror -eq ""){
        wl("$actionname : New-CMApplication -Name ""$PackageName"" -Description ""$description"" -Publisher ""$Publisher"" -SoftwareVersion ""$Version"" -LocalizedName ""$PackageName"" -LocalizedDescription ""$PackageName"" -OptionalReference ""$($tbUnikeyRef.text)"" -AutoInstall $true")
        try{
			New-CMApplication -Name "$PackageName" -Description "$description" -Publisher "$Publisher" -SoftwareVersion "$Version" -LocalizedName "$PackageName" -LocalizedDescription "$PackageName" -OptionalReference "$($tbUnikeyRef.text)" -AutoInstall $true -ErrorAction Stop
			wl("$actionname : created application $PackageName")
		}
		catch{
		    $returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Move Package"
    if($cbMovePackage.checked -And $returnerror -eq ""){
	    if($NewPackageLocation -ne ""){
            wl("$actionname : Move-CMObject -InputObject `$(Get-CMApplication -Name ""$PackageName"") -FolderPath ""$($SiteCode):\Application\$($NewPackageLocation)""")
            try{
    			Move-CMObject -InputObject $(Get-CMApplication -Name "$PackageName") -FolderPath "$($SiteCode):\Application\$($NewPackageLocation)" -ErrorAction Stop
    			wl("$actionname : moved application to $($NewPackageLocation)")
    		}
    		catch{
    		    $returnerror ="Error: "+ $_.Exception.Message
    		    wl($returnerror)
		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
    		}
	    }
    }
    $actionname = "Create Deployment Type"
    if($cbCreateDeploymentType.checked -And $returnerror -eq ""){
        if($rbMSI.checked){
            $MSIName = GetMSIFileName($defaultPackageLocation)
            if($MSIName -eq ""){
                $returnerror = ShowMessageBoxWithError("No MSI files were found in $defaultPackageLocation folder")
            }else{
                wl("$actionname : found msi file $MSIName")
                wl("$actionname : getting ProductCode property for $defaultPackageLocation\$($MSIName)")
                $tmp = GetProductCode("$defaultPackageLocation\$($MSIName)")
                $ProductCode = ($tmp|Out-String).Trim()
                if($ProductCode -like "*error*"){
                    $returnerror += $ProductCode
                }else{
                    wl("$actionname : ProductCode=$ProductCode")
                    $guidProductCode = [GUID]$ProductCode
                	$InstallCommand = "msiexec.exe /i ""$($MSIName)"""
                	If([System.IO.File]::Exists("$defaultPackageLocation\$($PackageName).mst")){$InstallCommand += " TRANSFORMS=$($PackageName).mst"}
                	$InstallCommand += " /qn /l* C:\Windows\Logs\$($PackageName)_I.log"
                	wl("$actionname : InstallCommand=$InstallCommand")
                	$UninstallCommand = "msiexec.exe /x $ProductCode /qn /l* C:\Windows\Logs\$($PackageName)_U.log"
                	wl("$actionname : UninstallCommand=$UninstallCommand")
                	wl("$actionname : Add-CMMsiDeploymentType -DeploymentTypeName `"$PackageName`" -InstallCommand `"$InstallCommand`" -ApplicationName `"$PackageName`" -ProductCode $ProductCode -ContentLocation `"$defaultPackageLocation`" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand `"$UninstallCommand`" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment `"$description`" -Force")
                    try{
                    	Add-CMMsiDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "$InstallCommand" -ApplicationName "$PackageName" -ProductCode $ProductCode -ContentLocation "$defaultPackageLocation\$($MSIName)" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "$UninstallCommand" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description" -Force -ErrorAction Stop
                    	wl("$actionname : created MSI deployment type")
                    }catch{
                    	$returnerror ="Error: "+ $_.Exception.Message
            		    wl($returnerror)
		                $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
                    }
                }
            }
        }
        if($rbPSAppDeploy.checked){
            if(CheckFileExists("$defaultPackageLocation\Deploy-Application.exe")){
                wl("$actionname : Add-CMScriptDeploymentType -DeploymentTypeName ""$PackageName"" -InstallCommand ""Deploy-Application.exe -DeploymentType """"Install"""" -DeployMode """"Silent"""""" -ApplicationName ""$PackageName"" -AddDetectionClause `$(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName ""$($DetectionKeyLocation)\$PackageName"" -PropertyType String -ValueName ""Installed"" -Is64Bit -Existence) -ContentLocation ""$defaultPackageLocation"" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand ""Deploy-Application.exe -DeploymentType """"Uninstall"""" -DeployMode """"Silent"""""" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment ""$description""")
                try{
                	Add-CMScriptDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "Deploy-Application.exe -DeploymentType ""Install"" -DeployMode ""Silent""" -ApplicationName "$PackageName" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName "$($DetectionKeyLocation)\$PackageName" -PropertyType String -ValueName "Installed" -Is64Bit -Existence) -ContentLocation "$defaultPackageLocation" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "Deploy-Application.exe -DeploymentType ""Uninstall"" -DeployMode ""Silent""" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description" -ErrorAction Stop
                	wl("$actionname : created PSAppdeploy scripted deployment type")
                }catch{
                	$returnerror ="Error: "+ $_.Exception.Message
        		    wl($returnerror)
    		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
                }
            }else{
                $returnerror = ShowMessageBoxWithError("Deploy-Application.exe file wasn't found in $defaultPackageLocation folder")
            }
        }
        if($rbBAT.checked){
            if((CheckFileExists("$defaultPackageLocation\install.bat")) -And (CheckFileExists("$defaultPackageLocation\uninstall.bat"))){
                wl("$actionname : Add-CMScriptDeploymentType -DeploymentTypeName ""$PackageName"" -InstallCommand ""install.bat"" -ApplicationName ""$PackageName"" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName ""$($DetectionKeyLocation)\$PackageName"" -PropertyType String -ValueName ""Installed"" -Is64Bit -Existence) -ContentLocation ""$defaultPackageLocation"" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand ""uninstall.bat"" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment ""$description""")
                try{
                	Add-CMScriptDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "install.bat" -ApplicationName "$PackageName" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName "$($DetectionKeyLocation)\$PackageName" -PropertyType String -ValueName "Installed" -Is64Bit -Existence) -ContentLocation "$defaultPackageLocation" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "uninstall.bat" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description" -ErrorAction Stop
                	wl("$actionname : created BAT scripted deployment type")
                }catch{
                	$returnerror ="Error: "+ $_.Exception.Message
        		    wl($returnerror)
    		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
                }
            }else{
                $returnerror = ShowMessageBoxWithError("Install.bat or uninstall.bat files were not found in $defaultPackageLocation folder")
            }
        }
    }
    $actionname = "Distribute Content"
    if($cbDistributeContent.checked -And $returnerror -eq ""){
        wl("$actionname : Start-CMContentDistribution -ApplicationName ""$PackageName"" -DisableContentDependencyDetection -DistributionPointGroupName $($DistributionPointGroups)")
        try{
        	Start-CMContentDistribution -ApplicationName "$PackageName" -DisableContentDependencyDetection -DistributionPointGroupName $($DistributionPointGroups.Split(';')) -ErrorAction Stop
        	wl("$actionname : distributed content to $DistributionPointGroups")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Create Install Device Collection"
    if($cbCreateInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : New-CMDeviceCollection -Name ""$InstallCollectionName"" -LimitingCollectionName ""$($LimitingDeviceCollectionName)"" -RefreshType Continuous -Comment ""$description""")
        try{
        	New-CMDeviceCollection -Name "$InstallCollectionName" -LimitingCollectionName "$($LimitingDeviceCollectionName)" -RefreshType Continuous -Comment "$description" -ErrorAction Stop
        	wl("$actionname : created Install Device Collection $InstallCollectionName")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Move Install Device Collection"
    if($cbMoveInstallDeviceCollection.checked -And $returnerror -eq "" -And $InstallDeviceCollectionLocation -ne ""){
        wl("$actionname : Move-CMObject -InputObject `$(Get-CMDeviceCollection -Name ""$InstallCollectionName"") -FolderPath ""$($SiteCode):\DeviceCollection\$($InstallDeviceCollectionLocation)""")
        try{
        	Move-CMObject -InputObject $(Get-CMDeviceCollection -Name "$InstallCollectionName") -FolderPath "$($SiteCode):\DeviceCollection\$($InstallDeviceCollectionLocation)" -ErrorAction Stop
        	wl("$actionname : moved it to $InstallDeviceCollectionLocation")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Create Uninstall Device Collection"
    if($cbCreateUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : New-CMDeviceCollection -Name ""$UninstallCollectionName"" -LimitingCollectionName ""$($LimitingDeviceCollectionName)"" -RefreshType Continuous -Comment ""$description""")
        try{
        	New-CMDeviceCollection -Name "$UninstallCollectionName" -LimitingCollectionName "$($LimitingDeviceCollectionName)" -RefreshType Continuous -Comment "$description" -ErrorAction Stop
        	wl("$actionname : created Uninstall Device Collection $UninstallCollectionName")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Move Uninstall Device Collection"
    if($cbMoveUninstallDeviceCollection.checked -And $returnerror -eq "" -And $UninstallDeviceCollectionLocation -ne ""){
        wl("$actionname : Move-CMObject -InputObject `$(Get-CMDeviceCollection -Name ""$UninstallCollectionName"") -FolderPath ""$($SiteCode):\DeviceCollection\$($UninstallDeviceCollectionLocation)""")
        try{
        	Move-CMObject -InputObject $(Get-CMDeviceCollection -Name "$UninstallCollectionName") -FolderPath "$($SiteCode):\DeviceCollection\$($UninstallDeviceCollectionLocation)" -ErrorAction Stop
        	wl("$actionname : moved it to $($UninstallDeviceCollectionLocation)")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Create Available User Collection"
    if($cbCreateAvailableUserCollection.checked -And $returnerror -eq ""){
        wl("$actionname : New-CMUserCollection -Name ""$UserCollectionName"" -LimitingCollectionName ""$($LimitingUserCollectionName)"" -RefreshType Continuous")
        try{
        	New-CMUserCollection -Name "$UserCollectionName" -LimitingCollectionName "$($LimitingUserCollectionName)" -RefreshType Continuous -ErrorAction Stop
        	wl("$actionname : created Install User Collection $UserCollectionName")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Move Available User Collection"
    if($cbMoveAvailableUserCollection.checked -And $returnerror -eq "" -And $UserCollectionLocation -ne ""){
        wl("$actionname : Move-CMObject -InputObject $(Get-CMUserCollection -Name ""$UserCollectionName"") -FolderPath ""$($SiteCode):\UserCollection\$($UserCollectionLocation)""")
        try{
        	Move-CMObject -InputObject $(Get-CMUserCollection -Name "$UserCollectionName") -FolderPath "$($SiteCode):\UserCollection\$($UserCollectionLocation)" -ErrorAction Stop
    	    wl("$actionname : moved it to $($UserCollectionLocation)")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Deploy To Install Device Collection (Required, Hidden)"
    if($cbDeployToInstallDeviceCollectionHiddenRequired.checked -And $returnerror -eq ""){
        wl("$actionname : New-CMApplicationDeployment -Name ""$PackageName"" -DeployAction Install -DeployPurpose Required -UserNotification HideAll -CollectionName ""$InstallCollectionName"" -Comment ""$description""")
        try{
        	New-CMApplicationDeployment -Name "$PackageName" -DeployAction Install -DeployPurpose Required -UserNotification HideAll -CollectionName "$InstallCollectionName" -Comment "$description" -ErrorAction Stop
        	wl("$actionname : assigned application for deployment to $InstallCollectionName")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Deploy To Uninstall Device Collection (Required, Hidden)"
    if($cbDeployToUninstallDeviceCollectionHiddenRequired.checked -And $returnerror -eq ""){
        wl("$actionname : New-CMApplicationDeployment -Name ""$PackageName"" -DeployAction Uninstall -DeployPurpose Required -UserNotification HideAll -CollectionName ""$UninstallCollectionName"" -Comment ""$description""")
        try{
        	New-CMApplicationDeployment -Name "$PackageName" -DeployAction Uninstall -DeployPurpose Required -UserNotification HideAll -CollectionName "$UninstallCollectionName" -Comment "$description" -ErrorAction Stop
        	wl("$actionname : assigned application for deployment to $UninstallCollectionName")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Deploy To User Collection (Available)"
    if($cbDeployToUserCollectionAvailable.checked -And $returnerror -eq ""){
        wl("$actionname : New-CMApplicationDeployment -Name ""$PackageName"" -DeployAction Install -DeployPurpose Available -UserNotification DisplayAll -CollectionName ""$UserCollectionName"" -Comment ""$description""")
        try{
        	New-CMApplicationDeployment -Name "$PackageName" -DeployAction Install -DeployPurpose Available -UserNotification DisplayAll -CollectionName "$UserCollectionName" -Comment "$description" -ErrorAction Stop
        	wl("$actionname : assigned application for deployment to $UserCollectionName")
        }catch{
        	$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Create Install AD Group (_I)"
    if($cbCreateInstallADGroup.checked -And $returnerror -eq ""){
        wl("$actionname : New-ADGroup ""$($ADGroupNamePrefix)$($InstallCollectionName)"" -Path ""$($ADOUPath)"" -GroupCategory Security -GroupScope Global -Description ""$localizedName"" -OtherAttributes @{info=""This group deploys $($tbUnikeyRef.text) """"$PackageName"""" to any Windows 10 devices specified - Please do not add users to the group. Devices should only be added with the necessary approval. If in doubt please liaise with the application owner.""}")
		try{
			New-ADGroup "$($ADGroupNamePrefix)$($InstallCollectionName)" -Path "$($ADOUPath)" -GroupCategory Security -GroupScope Global -Description "$localizedName" -OtherAttributes @{info="This group deploys $($tbUnikeyRef.text) `"$PackageName`" to any Windows 10 devices specified - Please do not add users to the group. Devices should only be added with the necessary approval. If in doubt please liaise with the application owner."} -ErrorAction Stop
        	wl("$actionname : created AD group $($ADGroupNamePrefix)$($InstallCollectionName)")
		}catch{
			$returnerror ="Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Create Uninstall AD Group (_U)"
    if($cbCreateUninstallADGroup.checked -And $returnerror -eq ""){
        wl("$actionname : New-ADGroup ""$($ADGroupNamePrefix)$($UninstallCollectionName)"" -Path ""$($ADOUPath)"" -GroupCategory Security -GroupScope Global -Description ""$localizedName""")
		try{
			New-ADGroup "$($ADGroupNamePrefix)$($UninstallCollectionName)" -Path "$($ADOUPath)" -GroupCategory Security -GroupScope Global -Description "$localizedName" -ErrorAction Stop
        	wl("$actionname : created AD group $($ADGroupNamePrefix)$($UninstallCollectionName)")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Create User Available AD Group (_A)"
    if($cbCreateUserAvailableADGroup.checked -And $returnerror -eq ""){
        wl("$actionname : New-ADGroup ""$($ADGroupNamePrefix)$($UserCollectionName)"" -Path ""$($ADOUPath)"" -GroupCategory Security -GroupScope Global -Description ""$localizedName"" -OtherAttributes @{info=""This group makes available $($tbUnikeyRef.text) """"$PackageName"""" for users specified - Please do not add computer to this group. Users should only be added with the necessary approval. If in doubt please liaise with the application owner.""}")
		try{
			New-ADGroup "$($ADGroupNamePrefix)$($UserCollectionName)" -Path "$($ADOUPath)h" -GroupCategory Security -GroupScope Global -Description "$localizedName" -OtherAttributes @{info="This group makes available $($tbUnikeyRef.text) `"$PackageName`" for users specified - Please do not add computer to this group. Users should only be added with the necessary approval. If in doubt please liaise with the application owner."} -ErrorAction Stop
        	wl("$actionname : created AD group $($ADGroupNamePrefix)$($UserCollectionName)")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Add AD Groups Query to Install Device Collection"
    if($cbAddADGroupQueryToInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Add-CMDeviceCollectionQueryMembershipRule -CollectionName ""$InstallCollectionName"" -RuleName ""$($ADGroupNamePrefix)$($InstallCollectionName)"" -QueryExpression ""$($ADGroupQuery)$($ADGroupNamePrefix)$($InstallCollectionName)'""")
		try{
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$InstallCollectionName" -RuleName "$($ADGroupNamePrefix)$($InstallCollectionName)" -QueryExpression "$($ADGroupQuery)$($ADGroupNamePrefix)$($InstallCollectionName)'" -ErrorAction Stop
        	wl("$actionname : created query for AD group $($ADGroupNamePrefix)$($InstallCollectionName) in $InstallCollectionName")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Add AD Groups Query to Uninstall Device Collection"
    if($cbAddADGroupQueryToUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Add-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UninstallCollectionName"" -RuleName ""$($ADGroupNamePrefix)$($UninstallCollectionName)"" -QueryExpression ""$($ADGroupQuery)$($ADGroupNamePrefix)$($UninstallCollectionName)'""")
		try{
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$UninstallCollectionName" -RuleName "$($ADGroupNamePrefix)$($UninstallCollectionName)" -QueryExpression "$($ADGroupQuery)$($ADGroupNamePrefix)$($UninstallCollectionName)'" -ErrorAction Stop
        	wl("$actionname : created query for AD group $($ADGroupNamePrefix)$($UninstallCollectionName) in $UninstallCollectionName")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Add AD Group Query To User Collection"
    if($cbAddADGroupQueryToUserCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Add-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UserCollectionName"" -RuleName ""$($ADGroupNamePrefix)$($UserCollectionName)"" -QueryExpression ""$($ADGroupUserQuery)$($ADGroupNamePrefix)$($UserCollectionName)'""")
		try{
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$UserCollectionName" -RuleName "$($ADGroupNamePrefix)$($UserCollectionName)" -QueryExpression "$($ADGroupUserQuery)$($ADGroupNamePrefix)$($UserCollectionName)'" -ErrorAction Stop
        	wl("$actionname : created query for AD group $($ADGroupNamePrefix)$($UserCollectionName) in $UserCollectionName")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Update Content On DPs"
    if($cbUpdateContentOnDPs.checked -And $returnerror -eq ""){
        wl("$actionname : Enumirating Deployment Types for $PackageName application")
		try{
            foreach($dt in $(Get-CMDeploymentType -ApplicationName "$PackageName")){
               wl("$actionname : Update-CMDistributionPoint -ApplicationName `"$PackageName`" -DeploymentTypeName `"$($dt.LocalizedDisplayName)`"")
               Update-CMDistributionPoint -ApplicationName "$PackageName" -DeploymentTypeName "$($dt.LocalizedDisplayName)" -ErrorAction Stop
        	   wl("$actionname : updated distributed content for $($dt.LocalizedDisplayName)")
            }            
			
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Request Machine Assignments"
    if($cbRequestMachineAssignments.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("$actionname : Start-Job -ScriptBlock {Invoke-Command -ComputerName ""$tm"" -ScriptBlock {Invoke-WmiMethod -Namespace ""Root\CCM"" -Class SMS_Client -Name TriggerSchedule -ArgumentList ""{00000000-0000-0000-0000-000000000021}""}}")
    		try{
    			Start-Job -ScriptBlock {Invoke-Command -ComputerName "$tm" -ScriptBlock {Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"}}
        	    wl("$actionname : started on $tm")
    		}catch{
    			$returnerror = "Error: "+ $_.Exception.Message
    		    wl($returnerror)
		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
    		}
        }
    }
    $actionname = "Add Test Machines To Install Device Collection"
    if($cbAddTestMachinesToInstallDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("$actionname : Add-CMDeviceCollectionDirectMembershipRule -CollectionName ""$InstallCollectionName"" -Resource `$(Get-CMDevice -Name $tm)")
    		try{
    			Add-CMDeviceCollectionDirectMembershipRule -CollectionName "$InstallCollectionName" -Resource $(Get-CMDevice -Name $tm) -ErrorAction Stop
        	    wl("$actionname : added $tm to $InstallCollectionName")
    		}catch{
    			$returnerror = "Error: "+ $_.Exception.Message
    		    wl($returnerror)
		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
    		}
        }
    }
    $actionname = "Remove Test Machines From Install Device Collection"
    if($cbRemoveTestMachinesFromInstallDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("$actionname : Remove-CMDeviceCollectionDirectMembershipRule -CollectionName ""$InstallCollectionName"" -ResourceName $tm -Force")
    		try{
    			Remove-CMDeviceCollectionDirectMembershipRule -CollectionName "$InstallCollectionName" -ResourceName $tm -Force -ErrorAction Stop
        	    wl("$actionname : removed $tm from $InstallCollectionName")
    		}catch{
    			$returnerror = "Error: "+ $_.Exception.Message
    		    wl($returnerror)
		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
    		}
        }
    }
    $actionname = "Add Test Machines To Uninstall Device Collection"
    if($cbAddTestMachinesToUninstallDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("$actionname : Add-CMDeviceCollectionDirectMembershipRule -CollectionName ""$UninstallCollectionName"" -Resource `$(Get-CMDevice -Name $tm)")
    		try{
    			Add-CMDeviceCollectionDirectMembershipRule -CollectionName "$UninstallCollectionName" -Resource $(Get-CMDevice -Name $tm) -ErrorAction Stop
        	    wl("$actionname : added $tm to $UninstallCollectionName")
    		}catch{
    			$returnerror = "Error: "+ $_.Exception.Message
    		    wl($returnerror)
		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
    		}
        }
    }
    $actionname = "Remove Test Machines From Uninstall Device Collection"
    if($cbRemoveTestMachinesFromUninstallDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("$actionname : Remove-CMDeviceCollectionDirectMembershipRule -CollectionName ""$UninstallCollectionName"" -ResourceName $tm -Force")
    		try{
    			Remove-CMDeviceCollectionDirectMembershipRule -CollectionName "$UninstallCollectionName" -ResourceName $tm -Force -ErrorAction Stop
        	    wl("$actionname : removed $tm from $UninstallCollectionName")
    		}catch{
    			$returnerror = "Error: "+ $_.Exception.Message
    		    wl($returnerror)
		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
    		}
        }
    }
    $actionname = "Add Test Users To Available User Collection"
    if($cbAddTestUsersToAvailableUserCollection.checked -And $returnerror -eq ""){
        foreach($tu in $testuser){
            wl("$actionname : Add-CMUserCollectionDirectMembershipRule -CollectionName ""$UserCollectionName"" -Resource `$(Get-CMUser -Name ""$($DomainPrefix)\$tu"")")
    		try{
    			Add-CMUserCollectionDirectMembershipRule -CollectionName "$UserCollectionName" -Resource $(Get-CMUser -Name "$($DomainPrefix)\$tu") -ErrorAction Stop
        	    wl("$actionname : added $tu to $InstallCollectionName")
    		}catch{
    			$returnerror = "Error: "+ $_.Exception.Message
    		    wl($returnerror)
		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
    		}
        }
    }
    $actionname = "Remove Test Users From User Collection"
    if($cbRemoveTestUsersFromAvailableUserCollection.checked -And $returnerror -eq ""){
        foreach($tu in $testuser){
            wl("$actionname : Remove-CMUserCollectionDirectMembershipRule -CollectionName ""$UserCollectionName"" -Resource `$(Get-CMUser -Name ""$($DomainPrefix)\$tu"") -Force")
    		try{
    			Remove-CMUserCollectionDirectMembershipRule -CollectionName "$UserCollectionName" -Resource $(Get-CMUser -Name "$($DomainPrefix)\$tu") -Force -ErrorAction Stop
        	    wl("$actionname : removed $tu from $UserCollectionName")
    		}catch{
    			$returnerror = "Error: "+ $_.Exception.Message
    		    wl($returnerror)
		        $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
    		}
        }
    }
    $actionname = "Remove Install Device Collection"
    if($cbRemoveInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-CMDeviceCollection -Name ""$InstallCollectionName"" -Force")
        try{
        	Remove-CMDeviceCollection -Name "$InstallCollectionName" -Force -ErrorAction Stop
        	wl("$actionname : Removed Install Device Collection $InstallCollectionName")
        }catch{
        	$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Remove Uninstall Device Collection"
    if($cbRemoveUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-CMDeviceCollection -Name ""$UninstallCollectionName"" -Force")
        try{
        	Remove-CMDeviceCollection -Name "$UninstallCollectionName" -Force -ErrorAction Stop
        	wl("$actionname : Removed Uninstall Device Collection $UninstallCollectionName")
        }catch{
        	$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Remove Available User Collection"
    if($cbRemoveAvailableUserCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-CMUserCollection -Name ""$UserCollectionName"" -Force")
        try{
        	Remove-CMUserCollection -Name "$UserCollectionName" -Force -ErrorAction Stop
        	wl("$actionname : Removed Install User Collection $UserCollectionName")
        }catch{
        	$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Remove Deployment To Install Device Collection"
    if($cbRemoveDeploymentToInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-CMApplicationDeployment -Name ""$PackageName"" -CollectionName ""$InstallCollectionName"" -Force")
        try{
        	Remove-CMApplicationDeployment -Name "$PackageName" -CollectionName "$InstallCollectionName" -Force -ErrorAction Stop
        	wl("$actionname : Removed Install Device Collection $InstallCollectionName")
        }catch{
        	$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Remove Deployment To Uninstall Device Collection"
    if($cbRemoveDeploymentToUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-CMApplicationDeployment -Name ""$PackageName"" -CollectionName ""$UninstallCollectionName"" -Force")
        try{
        	Remove-CMApplicationDeployment -Name "$PackageName" -CollectionName "$UninstallCollectionName" -Force -ErrorAction Stop
        	wl("$actionname : Removed Uninstall Device Collection $UninstallCollectionName")
        }catch{
        	$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Remove Deployment To Available User Collection"
    if($cbRemoveDeploymentToAvailableUserCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-CMApplicationDeployment -Name ""$PackageName"" -CollectionName ""$UserCollectionName"" -Force")
        try{
        	Remove-CMApplicationDeployment -Name "$PackageName" -CollectionName "$UserCollectionName" -Force -ErrorAction Stop
        	wl("$actionname : Removed Install User Collection $UserCollectionName")
        }catch{
        	$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
        }
    }
    $actionname = "Remove Install AD Group (_I)"
    if($cbRemoveInstallADGroup.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-ADGroup ""$($ADGroupNamePrefix)$($InstallCollectionName)"" -Confirm:$false")
		try{
			Remove-ADGroup "$($ADGroupNamePrefix)$($InstallCollectionName)" -Confirm:$false -ErrorAction Stop
        	wl("$actionname : Removed AD group $($ADGroupNamePrefix)$($InstallCollectionName)")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Remove Uninstall AD Group (_U)"
    if($cbRemoveUninstallADGroup.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-ADGroup ""$($ADGroupNamePrefix)$($UninstallCollectionName)"" -Confirm:$false")
		try{
			Remove-ADGroup "$($ADGroupNamePrefix)$($UninstallCollectionName)" -Confirm:$false -ErrorAction Stop
        	wl("$actionname : Removed AD group $($ADGroupNamePrefix)$($UninstallCollectionName)")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Remove User Available AD Group (_A)"
    if($cbRemoveUserAvailableADGroup.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-ADGroup ""$($ADGroupNamePrefix)$($UserCollectionName)"" -Confirm:$false")
		try{
			Remove-ADGroup "$($ADGroupNamePrefix)$($UserCollectionName)" -Confirm:$false -ErrorAction Stop
        	wl("$actionname : Removed AD group $($ADGroupNamePrefix)$($UserCollectionName)")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Remove AD Group Query to Install Device Collection"
    if($cbRemoveADGroupQueryToInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("$actionname : Remove-CMDeviceCollectionQueryMembershipRule -CollectionName ""$InstallCollectionName"" -RuleName ""$($ADGroupNamePrefix)$($InstallCollectionName)"" -Force")
		try{
			Remove-CMDeviceCollectionQueryMembershipRule -CollectionName "$InstallCollectionName" -RuleName "$($ADGroupNamePrefix)$($InstallCollectionName)" -Force -ErrorAction Stop
        	wl("$actionname : removed query for AD group $($ADGroupNamePrefix)$($InstallCollectionName) in $InstallCollectionName")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Remove AD Group Query to Uninstall Device Collection"
    if($cbRemoveADGroupQueryToUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Remove AD Group Query to Uninstall Device Collection: Remove-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UninstallCollectionName"" -RuleName ""$($ADGroupNamePrefix)$($UninstallCollectionName)"" -Force")
		try{
			Remove-CMDeviceCollectionQueryMembershipRule -CollectionName "$UninstallCollectionName" -RuleName "$($ADGroupNamePrefix)$($UninstallCollectionName)" -Force -ErrorAction Stop
        	wl("$actionname : removed query for AD group $($ADGroupNamePrefix)$($UninstallCollectionName) in $UninstallCollectionName")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    $actionname = "Remove AD Group Query To User Collection"
    if($cbRemoveADGroupQueryToUserCollection.checked -And $returnerror -eq ""){
        wl("Remove AD Group Query To User Collection: Remove-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UserCollectionName"" -RuleName ""$($ADGroupNamePrefix)$($UserCollectionName)"" -Force")
		try{
			Remove-CMDeviceCollectionQueryMembershipRule -CollectionName "$UserCollectionName" -RuleName "$($ADGroupNamePrefix)$($UserCollectionName)" -Force -ErrorAction Stop
        	wl("$actionname : removed query for AD group $($ADGroupNamePrefix)$($UserCollectionName) in $UserCollectionName")
		}catch{
			$returnerror = "Error: "+ $_.Exception.Message
		    wl($returnerror)
		    $returnerror = ShowMessageBoxWithError("$actionname returned $returnerror")
		}
    }
    if($returnerror -eq ""){
        $ddPresets.SelectedItem = "Not Set"
        PresetNotSet
        wl("Completed Executing Actions")
    }
}
 
 
 #### LoadApplicationsAndFolders
 
 
function RefreshApplicationFoldersClicked {
    wl("Refreshing application folders list...")
    $Instancekeys = (Get-WmiObject -Namespace "ROOT\SMS\Site_$($SiteCode)" -ComputerName $SiteServer `
         -Query "select DISTINCT ObjectPath from SMS_Applicationlatest where ModelName in (select InstanceKey from SMS_ObjectContainerItem where ObjectType=6000) ORDER BY ObjectPath").ObjectPath
    $ddApplicationFolders.Items.Clear
    $ddApplicationFolders.Items.Add("All Applications from all folders")
    foreach ($key in $Instancekeys){
        $ddApplicationFolders.Items.Add($key)
    }
    $ddApplicationFolders.SelectedItem = "All Applications from all folders" 
    wl("Populated $($Instancekeys.Count) folders")
    $btnRefreshApplications.enabled = $true
}

function RefreshApplicationsList { 
    $Instancekeys = @()
    wl("Refreshing applications list...")
    if($ddApplicationFolders.SelectedItem -eq "All Applications from all folders"){
        $Instancekeys = (Get-WmiObject -Namespace "ROOT\SMS\Site_$($SiteCode)" -ComputerName $SiteServer `
-Query "select LocalizedDisplayName from SMS_Applicationlatest ORDER BY LocalizedDisplayName").LocalizedDisplayName
    }else{
        $Instancekeys = (Get-WmiObject -Namespace "ROOT\SMS\Site_$($SiteCode)" -ComputerName $SiteServer `
-Query "select LocalizedDisplayName from SMS_Applicationlatest where ObjectPath='$($ddApplicationFolders.SelectedItem)' ORDER BY LocalizedDisplayName").LocalizedDisplayName
    }
    $ddApplications.Items.Clear()
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
        $app = Get-CMApplication $($ddApplications.SelectedItem) -ErrorAction Stop
        $tbPackageName.text = $ddApplications.SelectedItem
        $tbPublisher.text = $(if($app.Manufacturer -eq ""){""}else{$app.Manufacturer})
        $tbVersion.text = $(if($app.SoftwareVersion -eq ""){""}else{$app.SoftwareVersion})
        $tmppub = $(if($tbPublisher.text -eq ""){"_"}else{$tbPublisher.text})
        $tmpver = $(if($app.SoftwareVersion -eq ""){"_"}else{$($tbVersion.text).replace(".","")})
        $tbApplicationName.text = $($tbPackageName.text).replace("$($tmppub)_","").replace("$($tmpver)_","").replace("EN_01_W10_F","").replace("EN_02_W10_F","").replace("EN_03_W10_F","").replace("EN_04_W10_F","").replace("EN_05_W10_F","").replace("01_W10_F","").replace("02_W10_F","").replace("03_W10_F","").replace("04_W10_F","").replace("05_W10_F","").replace("_"," ").Trim()
        $tbUnikeyRef.text = ([XML]($app).SDMPackageXML).AppMgmtDigest.Application.CustomId.'#text'
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
 
 
 #### ClearForm
 
 
function ClearForm{
    ClearErrors
    $lbLog.Items.Clear
}

function ClearErrors{
    $errPackageName.visible = $false
    $errPublisher.visible = $false
    $errApplicationName.visible = $false
    $errVersion.visible = $false
}
    

#### Main


ClearForm
PresetNotSet


##### test - mode


$testMode = $false

function TestMode{

}

$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false


<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    FormSettings
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$FormSettings                    = New-Object system.Windows.Forms.Form
$FormSettings.ClientSize         = '620,840'
$FormSettings.text               = "Settings"
$FormSettings.TopMost            = $false

$LabelFormSettings1              = New-Object system.Windows.Forms.Label
$LabelFormSettings1.text         = "Default Test User"
$LabelFormSettings1.AutoSize     = $true
$LabelFormSettings1.width        = 25
$LabelFormSettings1.height       = 10
$LabelFormSettings1.location     = New-Object System.Drawing.Point(30,20)
$LabelFormSettings1.Font         = 'Microsoft Sans Serif,10'

$tbDefaultUserName               = New-Object system.Windows.Forms.TextBox
$tbDefaultUserName.multiline     = $false
$tbDefaultUserName.width         = 300
$tbDefaultUserName.height        = 20
$tbDefaultUserName.location      = New-Object System.Drawing.Point(300,15)
$tbDefaultUserName.Font          = 'Microsoft Sans Serif,10'

$LabelFormSettings2              = New-Object system.Windows.Forms.Label
$LabelFormSettings2.text         = "Default Test Machine"
$LabelFormSettings2.AutoSize     = $true
$LabelFormSettings2.width        = 25
$LabelFormSettings2.height       = 10
$LabelFormSettings2.location     = New-Object System.Drawing.Point(30,50)
$LabelFormSettings2.Font         = 'Microsoft Sans Serif,10'

$tbDefaultTestMachine            = New-Object system.Windows.Forms.TextBox
$tbDefaultTestMachine.multiline  = $false
$tbDefaultTestMachine.width      = 300
$tbDefaultTestMachine.height     = 20
$tbDefaultTestMachine.location   = New-Object System.Drawing.Point(300,45)
$tbDefaultTestMachine.Font       = 'Microsoft Sans Serif,10'

$LabelFormSettings3              = New-Object system.Windows.Forms.Label
$LabelFormSettings3.text         = "Site Code"
$LabelFormSettings3.AutoSize     = $true
$LabelFormSettings3.width        = 25
$LabelFormSettings3.height       = 10
$LabelFormSettings3.location     = New-Object System.Drawing.Point(30,80)
$LabelFormSettings3.Font         = 'Microsoft Sans Serif,10'

$LabelFormSettings4              = New-Object system.Windows.Forms.Label
$LabelFormSettings4.text         = "Site Server"
$LabelFormSettings4.AutoSize     = $true
$LabelFormSettings4.width        = 25
$LabelFormSettings4.height       = 10
$LabelFormSettings4.location     = New-Object System.Drawing.Point(30,110)
$LabelFormSettings4.Font         = 'Microsoft Sans Serif,10'

$tbSiteCode                      = New-Object system.Windows.Forms.TextBox
$tbSiteCode.multiline            = $false
$tbSiteCode.width                = 300
$tbSiteCode.height               = 20
$tbSiteCode.location             = New-Object System.Drawing.Point(300,75)
$tbSiteCode.Font                 = 'Microsoft Sans Serif,10'

$tbSiteServer                    = New-Object system.Windows.Forms.TextBox
$tbSiteServer.multiline          = $false
$tbSiteServer.width              = 300
$tbSiteServer.height             = 20
$tbSiteServer.location           = New-Object System.Drawing.Point(300,105)
$tbSiteServer.Font               = 'Microsoft Sans Serif,10'

$tbPackageRepository             = New-Object system.Windows.Forms.TextBox
$tbPackageRepository.multiline   = $false
$tbPackageRepository.width       = 300
$tbPackageRepository.height      = 20
$tbPackageRepository.location    = New-Object System.Drawing.Point(300,135)
$tbPackageRepository.Font        = 'Microsoft Sans Serif,10'

$LabelFormSettings5              = New-Object system.Windows.Forms.Label
$LabelFormSettings5.text         = "Package Repository"
$LabelFormSettings5.AutoSize     = $true
$LabelFormSettings5.width        = 25
$LabelFormSettings5.height       = 10
$LabelFormSettings5.location     = New-Object System.Drawing.Point(30,140)
$LabelFormSettings5.Font         = 'Microsoft Sans Serif,10'

$LabelFormSettings7              = New-Object system.Windows.Forms.Label
$LabelFormSettings7.text         = "Package Location in SCCM"
$LabelFormSettings7.AutoSize     = $true
$LabelFormSettings7.width        = 25
$LabelFormSettings7.height       = 10
$LabelFormSettings7.location     = New-Object System.Drawing.Point(30,200)
$LabelFormSettings7.Font         = 'Microsoft Sans Serif,10'

$tbNewPackageLocation            = New-Object system.Windows.Forms.TextBox
$tbNewPackageLocation.multiline  = $false
$tbNewPackageLocation.width      = 300
$tbNewPackageLocation.height     = 20
$tbNewPackageLocation.location   = New-Object System.Drawing.Point(300,195)
$tbNewPackageLocation.Font       = 'Microsoft Sans Serif,10'

$LabelFormSettings8              = New-Object system.Windows.Forms.Label
$LabelFormSettings8.text         = "Distribution Point Groups (; separated)"
$LabelFormSettings8.AutoSize     = $true
$LabelFormSettings8.width        = 25
$LabelFormSettings8.height       = 10
$LabelFormSettings8.location     = New-Object System.Drawing.Point(30,230)
$LabelFormSettings8.Font         = 'Microsoft Sans Serif,10'

$tbDistributionPointGroups       = New-Object system.Windows.Forms.TextBox
$tbDistributionPointGroups.multiline  = $false
$tbDistributionPointGroups.width  = 300
$tbDistributionPointGroups.height  = 20
$tbDistributionPointGroups.location  = New-Object System.Drawing.Point(300,225)
$tbDistributionPointGroups.Font  = 'Microsoft Sans Serif,10'

$LabelFormSettings9              = New-Object system.Windows.Forms.Label
$LabelFormSettings9.text         = "Install Device Collection Location"
$LabelFormSettings9.AutoSize     = $true
$LabelFormSettings9.width        = 25
$LabelFormSettings9.height       = 10
$LabelFormSettings9.location     = New-Object System.Drawing.Point(30,260)
$LabelFormSettings9.Font         = 'Microsoft Sans Serif,10'

$tbInstallDeviceCollectionLocation   = New-Object system.Windows.Forms.TextBox
$tbInstallDeviceCollectionLocation.multiline  = $false
$tbInstallDeviceCollectionLocation.width  = 300
$tbInstallDeviceCollectionLocation.height  = 20
$tbInstallDeviceCollectionLocation.location  = New-Object System.Drawing.Point(300,255)
$tbInstallDeviceCollectionLocation.Font  = 'Microsoft Sans Serif,10'

$LabelFormSettings10             = New-Object system.Windows.Forms.Label
$LabelFormSettings10.text        = "Uninstall Device Collection Location"
$LabelFormSettings10.AutoSize    = $true
$LabelFormSettings10.width       = 25
$LabelFormSettings10.height      = 10
$LabelFormSettings10.location    = New-Object System.Drawing.Point(30,290)
$LabelFormSettings10.Font        = 'Microsoft Sans Serif,10'

$tbUninstallDeviceCollectionLocation   = New-Object system.Windows.Forms.TextBox
$tbUninstallDeviceCollectionLocation.multiline  = $false
$tbUninstallDeviceCollectionLocation.width  = 300
$tbUninstallDeviceCollectionLocation.height  = 20
$tbUninstallDeviceCollectionLocation.location  = New-Object System.Drawing.Point(300,285)
$tbUninstallDeviceCollectionLocation.Font  = 'Microsoft Sans Serif,10'

$LabelFormSettings11             = New-Object system.Windows.Forms.Label
$LabelFormSettings11.text        = "LimitingDeviceCollectionName"
$LabelFormSettings11.AutoSize    = $true
$LabelFormSettings11.width       = 25
$LabelFormSettings11.height      = 10
$LabelFormSettings11.location    = New-Object System.Drawing.Point(30,320)
$LabelFormSettings11.Font        = 'Microsoft Sans Serif,10'

$tbLimitingDeviceCollectionName   = New-Object system.Windows.Forms.TextBox
$tbLimitingDeviceCollectionName.multiline  = $false
$tbLimitingDeviceCollectionName.width  = 300
$tbLimitingDeviceCollectionName.height  = 20
$tbLimitingDeviceCollectionName.location  = New-Object System.Drawing.Point(300,315)
$tbLimitingDeviceCollectionName.Font  = 'Microsoft Sans Serif,10'

$LabelFormSettings12             = New-Object system.Windows.Forms.Label
$LabelFormSettings12.text        = "Limiting User Collection Name"
$LabelFormSettings12.AutoSize    = $true
$LabelFormSettings12.width       = 25
$LabelFormSettings12.height      = 10
$LabelFormSettings12.location    = New-Object System.Drawing.Point(30,350)
$LabelFormSettings12.Font        = 'Microsoft Sans Serif,10'

$tbLimitingUserCollectionName    = New-Object system.Windows.Forms.TextBox
$tbLimitingUserCollectionName.multiline  = $false
$tbLimitingUserCollectionName.width  = 300
$tbLimitingUserCollectionName.height  = 20
$tbLimitingUserCollectionName.location  = New-Object System.Drawing.Point(300,345)
$tbLimitingUserCollectionName.Font  = 'Microsoft Sans Serif,10'

$LabelFormSettings13             = New-Object system.Windows.Forms.Label
$LabelFormSettings13.text        = "User Collection Location"
$LabelFormSettings13.AutoSize    = $true
$LabelFormSettings13.width       = 25
$LabelFormSettings13.height      = 10
$LabelFormSettings13.location    = New-Object System.Drawing.Point(30,380)
$LabelFormSettings13.Font        = 'Microsoft Sans Serif,10'

$tbUserCollectionLocation        = New-Object system.Windows.Forms.TextBox
$tbUserCollectionLocation.multiline  = $false
$tbUserCollectionLocation.width  = 300
$tbUserCollectionLocation.height  = 20
$tbUserCollectionLocation.location  = New-Object System.Drawing.Point(300,375)
$tbUserCollectionLocation.Font   = 'Microsoft Sans Serif,10'

$LabelFormSettings14             = New-Object system.Windows.Forms.Label
$LabelFormSettings14.text        = "ADOUPath (Active Directory Folder "
$LabelFormSettings14.AutoSize    = $true
$LabelFormSettings14.width       = 25
$LabelFormSettings14.height      = 10
$LabelFormSettings14.location    = New-Object System.Drawing.Point(30,410)
$LabelFormSettings14.Font        = 'Microsoft Sans Serif,10'

$tbADOUPath                      = New-Object system.Windows.Forms.TextBox
$tbADOUPath.multiline            = $false
$tbADOUPath.width                = 300
$tbADOUPath.height               = 20
$tbADOUPath.location             = New-Object System.Drawing.Point(300,405)
$tbADOUPath.Font                 = 'Microsoft Sans Serif,10'

$LabelFormSettings15             = New-Object system.Windows.Forms.Label
$LabelFormSettings15.text        = "DomainPrefix"
$LabelFormSettings15.AutoSize    = $true
$LabelFormSettings15.width       = 25
$LabelFormSettings15.height      = 10
$LabelFormSettings15.location    = New-Object System.Drawing.Point(30,450)
$LabelFormSettings15.Font        = 'Microsoft Sans Serif,10'

$tbDomainPrefix                  = New-Object system.Windows.Forms.TextBox
$tbDomainPrefix.multiline        = $false
$tbDomainPrefix.width            = 300
$tbDomainPrefix.height           = 20
$tbDomainPrefix.location         = New-Object System.Drawing.Point(300,445)
$tbDomainPrefix.Font             = 'Microsoft Sans Serif,10'

$LabelFormSettings16             = New-Object system.Windows.Forms.Label
$LabelFormSettings16.text        = "AD Group Machine Query"
$LabelFormSettings16.AutoSize    = $true
$LabelFormSettings16.width       = 25
$LabelFormSettings16.height      = 10
$LabelFormSettings16.location    = New-Object System.Drawing.Point(30,480)
$LabelFormSettings16.Font        = 'Microsoft Sans Serif,10'

$tbADGroupQuery                  = New-Object system.Windows.Forms.TextBox
$tbADGroupQuery.multiline        = $false
$tbADGroupQuery.width            = 300
$tbADGroupQuery.height           = 20
$tbADGroupQuery.location         = New-Object System.Drawing.Point(300,475)
$tbADGroupQuery.Font             = 'Microsoft Sans Serif,10'

$LabelFormSettings17             = New-Object system.Windows.Forms.Label
$LabelFormSettings17.text        = "AD Group User Query"
$LabelFormSettings17.AutoSize    = $true
$LabelFormSettings17.width       = 25
$LabelFormSettings17.height      = 10
$LabelFormSettings17.location    = New-Object System.Drawing.Point(30,510)
$LabelFormSettings17.Font        = 'Microsoft Sans Serif,10'

$tbADGroupUserQuery              = New-Object system.Windows.Forms.TextBox
$tbADGroupUserQuery.multiline    = $false
$tbADGroupUserQuery.width        = 300
$tbADGroupUserQuery.height       = 20
$tbADGroupUserQuery.location     = New-Object System.Drawing.Point(300,505)
$tbADGroupUserQuery.Font         = 'Microsoft Sans Serif,10'

$LabelFormSettings18             = New-Object system.Windows.Forms.Label
$LabelFormSettings18.text        = "Detection Key Location"
$LabelFormSettings18.AutoSize    = $true
$LabelFormSettings18.width       = 25
$LabelFormSettings18.height      = 10
$LabelFormSettings18.location    = New-Object System.Drawing.Point(30,540)
$LabelFormSettings18.Font        = 'Microsoft Sans Serif,10'

$tbDetectionKeyLocation          = New-Object system.Windows.Forms.TextBox
$tbDetectionKeyLocation.multiline  = $false
$tbDetectionKeyLocation.width    = 300
$tbDetectionKeyLocation.height   = 20
$tbDetectionKeyLocation.location  = New-Object System.Drawing.Point(300,535)
$tbDetectionKeyLocation.Font     = 'Microsoft Sans Serif,10'

$LabelFormSettings19             = New-Object system.Windows.Forms.Label
$LabelFormSettings19.text        = "Dropdown Application Preselected Folder"
$LabelFormSettings19.AutoSize    = $true
$LabelFormSettings19.width       = 240
$LabelFormSettings19.height      = 10
$LabelFormSettings19.location    = New-Object System.Drawing.Point(5,30)
$LabelFormSettings19.Font        = 'Microsoft Sans Serif,10'

$tbSelectedApplicationFolder     = New-Object system.Windows.Forms.TextBox
$tbSelectedApplicationFolder.multiline  = $false
$tbSelectedApplicationFolder.width  = 300
$tbSelectedApplicationFolder.height  = 20
$tbSelectedApplicationFolder.location  = New-Object System.Drawing.Point(270,25)
$tbSelectedApplicationFolder.Font  = 'Microsoft Sans Serif,10'

$cbPreloadApplicationFolders     = New-Object system.Windows.Forms.CheckBox
$cbPreloadApplicationFolders.text  = "Preload All Application Folders"
$cbPreloadApplicationFolders.AutoSize  = $true
$cbPreloadApplicationFolders.width  = 95
$cbPreloadApplicationFolders.height  = 20
$cbPreloadApplicationFolders.location  = New-Object System.Drawing.Point(20,60)
$cbPreloadApplicationFolders.Font  = 'Microsoft Sans Serif,10'

$cbOnlyPreselectFolder           = New-Object system.Windows.Forms.CheckBox
$cbOnlyPreselectFolder.text      = "Preload Only pre-selected folder"
$cbOnlyPreselectFolder.AutoSize  = $true
$cbOnlyPreselectFolder.width     = 95
$cbOnlyPreselectFolder.height    = 20
$cbOnlyPreselectFolder.location  = New-Object System.Drawing.Point(20,90)
$cbOnlyPreselectFolder.Font      = 'Microsoft Sans Serif,10'

$cbLoadApplications              = New-Object system.Windows.Forms.CheckBox
$cbLoadApplications.text         = "Pre-load Applications List for Selected Folder"
$cbLoadApplications.AutoSize     = $true
$cbLoadApplications.width        = 95
$cbLoadApplications.height       = 20
$cbLoadApplications.location     = New-Object System.Drawing.Point(20,120)
$cbLoadApplications.Font         = 'Microsoft Sans Serif,10'

$LabelFormSettings20             = New-Object system.Windows.Forms.Label
$LabelFormSettings20.text        = "where AD Group will be created)"
$LabelFormSettings20.AutoSize    = $true
$LabelFormSettings20.width       = 25
$LabelFormSettings20.height      = 10
$LabelFormSettings20.location    = New-Object System.Drawing.Point(71,427)
$LabelFormSettings20.Font        = 'Microsoft Sans Serif,10'

$GroupboxFormSettings1           = New-Object system.Windows.Forms.Groupbox
$GroupboxFormSettings1.height    = 150
$GroupboxFormSettings1.width     = 580
$GroupboxFormSettings1.text      = "Actions on script execution"
$GroupboxFormSettings1.location  = New-Object System.Drawing.Point(20,610)

$btnFormSettingsClose            = New-Object system.Windows.Forms.Button
$btnFormSettingsClose.text       = "Close"
$btnFormSettingsClose.width      = 60
$btnFormSettingsClose.height     = 30
$btnFormSettingsClose.location   = New-Object System.Drawing.Point(550,780)
$btnFormSettingsClose.Font       = 'Microsoft Sans Serif,10'

$btnFormSettingsSave             = New-Object system.Windows.Forms.Button
$btnFormSettingsSave.text        = "Save"
$btnFormSettingsSave.width       = 60
$btnFormSettingsSave.height      = 30
$btnFormSettingsSave.location    = New-Object System.Drawing.Point(30,780)
$btnFormSettingsSave.Font        = 'Microsoft Sans Serif,10'

$lblFormSettingsMessage          = New-Object system.Windows.Forms.Label
$lblFormSettingsMessage.text     = "Saved Successfully"
$lblFormSettingsMessage.AutoSize  = $true
$lblFormSettingsMessage.width    = 25
$lblFormSettingsMessage.height   = 10
$lblFormSettingsMessage.location  = New-Object System.Drawing.Point(240,790)
$lblFormSettingsMessage.Font     = 'Microsoft Sans Serif,10'
$lblFormSettingsMessage.ForeColor  = "#7ed321"

$btnFormSettingsLoad             = New-Object system.Windows.Forms.Button
$btnFormSettingsLoad.text        = "Load"
$btnFormSettingsLoad.width       = 60
$btnFormSettingsLoad.height      = 30
$btnFormSettingsLoad.location    = New-Object System.Drawing.Point(120,780)
$btnFormSettingsLoad.Font        = 'Microsoft Sans Serif,10'

$btnFormSettingsDefaults         = New-Object system.Windows.Forms.Button
$btnFormSettingsDefaults.text    = "Defaults"
$btnFormSettingsDefaults.width   = 70
$btnFormSettingsDefaults.height  = 30
$btnFormSettingsDefaults.location  = New-Object System.Drawing.Point(420,780)
$btnFormSettingsDefaults.Font    = 'Microsoft Sans Serif,10'

$tbADGroupNamePrefix             = New-Object system.Windows.Forms.TextBox
$tbADGroupNamePrefix.multiline   = $false
$tbADGroupNamePrefix.width       = 300
$tbADGroupNamePrefix.height      = 20
$tbADGroupNamePrefix.location    = New-Object System.Drawing.Point(300,565)
$tbADGroupNamePrefix.Font        = 'Microsoft Sans Serif,10'

$LabelFormSettings21             = New-Object system.Windows.Forms.Label
$LabelFormSettings21.text        = "ADGroupNamePrefix"
$LabelFormSettings21.AutoSize    = $true
$LabelFormSettings21.width       = 25
$LabelFormSettings21.height      = 10
$LabelFormSettings21.location    = New-Object System.Drawing.Point(30,570)
$LabelFormSettings21.Font        = 'Microsoft Sans Serif,10'

$FormSettingsCurrentMachine      = New-Object system.Windows.Forms.Button
$FormSettingsCurrentMachine.text  = "Current Machine"
$FormSettingsCurrentMachine.width  = 110
$FormSettingsCurrentMachine.height  = 30
$FormSettingsCurrentMachine.location  = New-Object System.Drawing.Point(180,40)
$FormSettingsCurrentMachine.Font  = 'Microsoft Sans Serif,10'

$FormSettingsCurrentUser         = New-Object system.Windows.Forms.Button
$FormSettingsCurrentUser.text    = "Current User"
$FormSettingsCurrentUser.width   = 110
$FormSettingsCurrentUser.height  = 30
$FormSettingsCurrentUser.location  = New-Object System.Drawing.Point(180,5)
$FormSettingsCurrentUser.Font    = 'Microsoft Sans Serif,10'

$FormSettings.controls.AddRange(@($LabelFormSettings1,$tbDefaultUserName,$LabelFormSettings2,$tbDefaultTestMachine,$LabelFormSettings3,$LabelFormSettings4,$tbSiteCode,$tbSiteServer,$tbPackageRepository,$LabelFormSettings5,$LabelFormSettings7,$tbNewPackageLocation,$LabelFormSettings8,$tbDistributionPointGroups,$LabelFormSettings9,$tbInstallDeviceCollectionLocation,$LabelFormSettings10,$tbUninstallDeviceCollectionLocation,$LabelFormSettings11,$tbLimitingDeviceCollectionName,$LabelFormSettings12,$tbLimitingUserCollectionName,$LabelFormSettings13,$tbUserCollectionLocation,$LabelFormSettings14,$tbADOUPath,$LabelFormSettings15,$tbDomainPrefix,$LabelFormSettings16,$tbADGroupQuery,$LabelFormSettings17,$tbADGroupUserQuery,$LabelFormSettings18,$tbDetectionKeyLocation,$LabelFormSettings20,$GroupboxFormSettings1,$btnFormSettingsClose,$btnFormSettingsSave,$lblFormSettingsMessage,$btnFormSettingsLoad,$btnFormSettingsDefaults,$tbADGroupNamePrefix,$LabelFormSettings21,$FormSettingsCurrentMachine,$FormSettingsCurrentUser))
$GroupboxFormSettings1.controls.AddRange(@($LabelFormSettings19,$tbSelectedApplicationFolder,$cbPreloadApplicationFolders,$cbOnlyPreselectFolder,$cbLoadApplications))

$btnFormSettingsClose.Add_Click({ FormSettingsbtnCloseClicked })
$btnFormSettingsSave.Add_Click({ FormSettingsbtnSaveClicked })
$btnFormSettingsDefaults.Add_Click({ btnFormSettingsDefaultsClicked })
$btnFormSettingsLoad.Add_Click({ btnFormSettingsLoadClicked })
$FormSettings.Add_Shown({ FormSettingsShown })
$FormSettingsCurrentUser.Add_Click({ FormSettingsCurrentUserClicked })
$FormSettingsCurrentMachine.Add_Click({ FormSettingsCurrentMachineClicked })
$FormSettings.Add_Closing({param($sender,$e)
	FormSettingsbtnCloseClicked
    $e.Cancel= $true
})

$FormSettings.FormBorderStyle = 'Fixed3D'
$FormSettings.MaximizeBox = $false


function btnFormSettingsDefaultsClicked { 
    SetDefaultVariables
    $tbDefaultUserName.text = $tbTestUser.text
    $tbDefaultTestMachine.text = $tbTestMachine.text
    $tbSiteCode.text = $SiteCode
    $tbSiteServer.text = $SiteServer
    $tbPackageRepository.text = $PackageRepository
   # $tbLogFileLocation.text = $LogFileLocation.text
    $tbNewPackageLocation.text = $NewPackageLocation
    $tbDistributionPointGroups.text = $DistributionPointGroups
    $tbInstallDeviceCollectionLocation.text = $InstallDeviceCollectionLocation
    $tbUninstallDeviceCollectionLocation.text = $UninstallDeviceCollectionLocation
    $tbLimitingDeviceCollectionName.text = $LimitingDeviceCollectionName
    $tbLimitingUserCollectionName.text = $LimitingUserCollectionName
    $tbUserCollectionLocation.text = $UserCollectionLocation
    $tbADOUPath.text = $ADOUPath
    $tbDomainPrefix.text = $DomainPrefix
    $tbADGroupQuery.text = $ADGroupQuery
    $tbADGroupUserQuery.text = $ADGroupUserQuery
    $tbDetectionKeyLocation.text = $DetectionKeyLocation
    $tbADGroupNamePrefix.text = $ADGroupNamePrefix
    $tbSelectedApplicationFolder.text = $ApplicationFoldersSelectedItem
    $cbPreloadApplicationFolders.checked = ConvertString01ToTrueFalse($PreloadApplicationFolders)
    $cbOnlyPreselectFolder.checked = ConvertString01ToTrueFalse($OnlyPreselectFolder)
    $cbLoadApplications.checked = ConvertString01ToTrueFalse($LoadApplications)
    
}
function btnFormSettingsLoadClicked { 
    if([System.IO.File]::Exists("$PSScriptRoot\Settings.xml")){
		cd C:\
        $xmlread = [xml](Get-Content "$PSScriptRoot\Settings.xml")
        $tbDefaultUserName.text = $xmlread.Settings.General.DefaultUserName
        $tbDefaultTestMachine.text = $xmlread.Settings.General.DefaultTestMachine
        $tbSiteCode.text = $xmlread.Settings.General.SiteCode
        $tbSiteServer.text = $xmlread.Settings.General.SiteServer
        $tbPackageRepository.text = $xmlread.Settings.General.PackageRepository
       # $tbLogFileLocation.text = $xmlread.Settings.General.LogFileLocation
        $tbNewPackageLocation.text = $xmlread.Settings.General.NewPackageLocation
        $tbDistributionPointGroups.text = $xmlread.Settings.General.DistributionPointGroups
        $tbInstallDeviceCollectionLocation.text = $xmlread.Settings.General.InstallDeviceCollectionLocation
        $tbUninstallDeviceCollectionLocation.text = $xmlread.Settings.General.UninstallDeviceCollectionLocation
        $tbLimitingDeviceCollectionName.text = $xmlread.Settings.General.LimitingDeviceCollectionName
        $tbLimitingUserCollectionName.text = $xmlread.Settings.General.LimitingUserCollectionName
        $tbUserCollectionLocation.text = $xmlread.Settings.General.UserCollectionLocation
        $tbADOUPath.text = $xmlread.Settings.General.ADOUPath
        $tbDomainPrefix.text = $xmlread.Settings.General.DomainPrefix
        $tbADGroupQuery.text = $xmlread.Settings.General.ADGroupQuery
        $tbADGroupUserQuery.text = $xmlread.Settings.General.ADGroupUserQuery
        $tbDetectionKeyLocation.text = $xmlread.Settings.General.DetectionKeyLocation
        $tbSelectedApplicationFolder.text = $xmlread.Settings.ActionsOnScriptExecution.SelectedApplicationFolder
        $cbPreloadApplicationFolders.checked = ConvertString01ToTrueFalse($xmlread.Settings.ActionsOnScriptExecution.PreloadApplicationFolders)
        $cbOnlyPreselectFolder.checked = ConvertString01ToTrueFalse($xmlread.Settings.ActionsOnScriptExecution.OnlyPreselectFolder)
        $cbLoadApplications.checked = ConvertString01ToTrueFalse($xmlread.Settings.ActionsOnScriptExecution.LoadApplications)
    }else{
        $lblFormSettingsMessage.text = "Settings File Doesn't Exist"
        $lblFormSettingsMessage.visible = $true
    }
}
function FormSettingsbtnSaveClicked {
    if([System.IO.File]::Exists("$PSScriptRoot\Settings.xml")){
        Remove-Item -Path "$PSScriptRoot\Settings.xml" -Force
    }
    $xmlsettings = New-Object System.Xml.XmlWriterSettings
    $xmlsettings.Indent = $true
    $xmlsettings.IndentChars = "    "
    $XmlWriter = [System.XML.XmlWriter]::Create("$PSScriptRoot\Settings.xml", $xmlsettings)
    
    $xmlWriter.WriteStartDocument()
    $xmlWriter.WriteProcessingInstruction("xml-stylesheet", "type='text/xsl' href='style.xsl'")
    
    $xmlWriter.WriteStartElement("Settings")  
        $xmlWriter.WriteStartElement("General")
            $xmlWriter.WriteElementString("DefaultUserName",$tbDefaultUserName.text)
            $xmlWriter.WriteElementString("DefaultTestMachine",$tbDefaultTestMachine.text)
            $xmlWriter.WriteElementString("SiteCode",$tbSiteCode.text)
            $xmlWriter.WriteElementString("SiteServer",$tbSiteServer.text)
            $xmlWriter.WriteElementString("PackageRepository",$tbPackageRepository.text)
           # $xmlWriter.WriteElementString("LogFileLocation",$tbLogFileLocation.text)
            $xmlWriter.WriteElementString("NewPackageLocation",$tbNewPackageLocation.text)
            $xmlWriter.WriteElementString("DistributionPointGroups",$tbDistributionPointGroups.text)
            $xmlWriter.WriteElementString("InstallDeviceCollectionLocation",$tbInstallDeviceCollectionLocation.text)
            $xmlWriter.WriteElementString("UninstallDeviceCollectionLocation",$tbUninstallDeviceCollectionLocation.text)
            $xmlWriter.WriteElementString("LimitingDeviceCollectionName",$tbLimitingDeviceCollectionName.text)
            $xmlWriter.WriteElementString("LimitingUserCollectionName",$tbLimitingUserCollectionName.text)
            $xmlWriter.WriteElementString("UserCollectionLocation",$tbUserCollectionLocation.text)
            $xmlWriter.WriteElementString("ADOUPath",$tbADOUPath.text)
            $xmlWriter.WriteElementString("DomainPrefix",$tbDomainPrefix.text)
            $xmlWriter.WriteElementString("ADGroupQuery",$tbADGroupQuery.text)
            $xmlWriter.WriteElementString("ADGroupUserQuery",$tbADGroupUserQuery.text)
    		$xmlWriter.WriteElementString("DetectionKeyLocation",$tbDetectionKeyLocation.text)
        $xmlWriter.WriteEndElement() 
        $xmlWriter.WriteStartElement("ActionsOnScriptExecution")
            $xmlWriter.WriteElementString("SelectedApplicationFolder",$tbSelectedApplicationFolder.text)
            $xmlWriter.WriteElementString("PreloadApplicationFolders",$(ConvertTrueFalseToString01($cbPreloadApplicationFolders.checked)))
            $xmlWriter.WriteElementString("OnlyPreselectFolder",$(ConvertTrueFalseToString01($cbOnlyPreselectFolder.checked)))
            $xmlWriter.WriteElementString("LoadApplications",$(ConvertTrueFalseToString01($cbLoadApplications.checked)))
        $xmlWriter.WriteEndElement() 
    $xmlWriter.WriteEndElement()
    
    $xmlWriter.WriteEndDocument()
    $xmlWriter.Flush()
    $xmlWriter.Close()
    $lblFormSettingsMessage.text = "Saved successfully"
    $lblFormSettingsMessage.visible = $true
    SetVariables
}


function FormSettingsbtnCloseClicked {
    SetVariables
    $FormSettings.Visible = $false
}


function ConvertTrueFalseToString01($boolval){
    $retval="0"
    if($boolval){
        $retval="1"
    }
    return $retval
}


function ConvertString01ToTrueFalse($strval){
    $retval=$false
    if($strval -eq "1"){
        $retval=$true
    }
    return $retval
}


function FormSettingsShown { 
    btnFormSettingsDefaultsClicked
    btnFormSettingsLoadClicked
    $lblFormSettingsMessage.text = ""
    $lblFormSettingsMessage.visible = $false
}


function FormSettingsCurrentMachineClicked { 
    $tbDefaultTestMachine.text = "$env:ComputerName"
}
function FormSettingsCurrentUserClicked { 
    $tbDefaultUserName.text = "$env:UserName"
}


 

function btnValidatePackageStateClicked { 
    $appname = $tbPackageName.text
    #To avoid wildcards
    $appname = $appname.Replace("*","")
    if(CheckApplicationExists($appname)){
        wl("ValidatePackageState: $appname application exists")
        $action = "Application DeploymentType"
        $retval = $true
        try{
            $tmpObj = Get-CMDeploymentType -ApplicationName $appname -DeploymentTypeName $appname -ErrorAction Stop
            if($tmpObj -eq $null){
                wl("Error: $action $appname wasn't found")
                $retval = $false
            }
        }
        catch{
            wl($_.Exception.Message)
            $retval = $false
        }
        finally{
            if($retval){
                wl("ValidatePackageState: $action $appname exists")
            }
        }
        $action = "Installation Collection"
        if($true){
            $retval = $true
            try{
                $tmpObj = Get-CMDeviceCollection -Name "$($appname)_I" -ErrorAction Stop
                if($tmpObj -eq $null){
                    wl("Error: $action $($appname)_I wasn't found")
                    $retval = $false
                }
            }
            catch{
                wl($_.Exception.Message)
                $retval = $false
            }
            finally{
                if($retval){
                    wl("ValidatePackageState: $action $($appname)_I exists")
                }
            }
        }
        $action = "Uninstallation Collection"
        if($true){
            $retval = $true
            try{
                $tmpObj = Get-CMDeviceCollection -Name "$($appname)_U" -ErrorAction Stop
                if($tmpObj -eq $null){
                    wl("Error: Application Uninstallation Collection $($appname)_U wasn't found")
                    $retval = $false
                }
            }
            catch{
                wl($_.Exception.Message)
                $retval = $false
            }
            finally{
                if($retval){
                    wl("ValidatePackageState: $action $($appname)_I exists")
                }
            }
        }
        $action = "Deployment for collection"
        if($true){
            $retval = $true
            try{
                $tmpObj = Get-CMApplicationDeployment -Name $appname -CollectionName "$($appname)_I" -ErrorAction Stop
                if($tmpObj -eq $null){
                    wl("Error: $action $($appname)_I wasn't found")
                    $retval = $false
                }
            }
            catch{
                wl($_.Exception.Message)
                $retval = $false
            }
            finally{
                if($retval){
                    wl("ValidatePackageState: $action $($appname)_I exists")
                }
            }
        }
        if($true){
            $retval = $true
            try{
                $tmpObj = Get-CMApplicationDeployment -Name $appname -CollectionName "$($appname)_U" -ErrorAction Stop
                if($tmpObj -eq $null){
                    wl("Error: $action $($appname)_U wasn't found")
                    $retval = $false
                }
            }
            catch{
                wl($_.Exception.Message)
                $retval = $false
            }
            finally{
                if($retval){
                    wl("ValidatePackageState: $action $($appname)_U exists")
                }
            }
        }
        $action = "Install ADGroup"
        if($true){
            $retval = $true
            try{
                $tmpObj = Get-ADGroup "SCCM_W10_APP_$($appname)_I" -ErrorAction Stop
                if($tmpObj -eq $null){
                    wl("Error: $action SCCM_W10_APP_$($appname)_I wasn't found")
                    $retval = $false
                }
            }
            catch{
                wl($_.Exception.Message)
                $retval = $false
            }
            finally{
                if($retval){
                    wl("ValidatePackageState: $action SCCM_W10_APP_$($appname)_I exists")
                }
            }
        }
        $action = "Uninstall ADGroup"
        if($true){
            $retval = $true
            try{
                $tmpObj = Get-ADGroup "SCCM_W10_APP_$($appname)_U" -ErrorAction Stop
                if($tmpObj -eq $null){
                    wl("Error: $action SCCM_W10_APP_$($appname)_U wasn't found")
                    $retval = $false
                }
            }
            catch{
                wl($_.Exception.Message)
                $retval = $false
            }
            finally{
                if($retval){
                    wl("ValidatePackageState: $action SCCM_W10_APP_$($appname)_U exists")
                }
            }
        }
        $action = "Install Collection AD Group Query"
        if($true){
            $retval = $true
            try{
                $tmpObj = Get-CMCollectionQueryMembershipRule -CollectionName "$($appname)_I" -RuleName "SCCM_W10_APP_$($appname)_I" -ErrorAction Stop
                if($tmpObj -eq $null){
                    wl("Error: $action SCCM_W10_APP_$($appname)_I wasn't found")
                    $retval = $false
                }
            }
            catch{
                wl($_.Exception.Message)
                $retval = $false
            }
            finally{
                if($retval){
                    wl("ValidatePackageState: $action SCCM_W10_APP_$($appname)_I exists")
                }
            }
        }
        $action = "Uninstall Collection AD Group Query"
        if($true){
            $retval = $true
            try{
                $tmpObj = Get-CMCollectionQueryMembershipRule -CollectionName "$($appname)_U" -RuleName "SCCM_W10_APP_$($appname)_U" -ErrorAction Stop
                if($tmpObj -eq $null){
                    wl("Error: $action SCCM_W10_APP_$($appname)_U wasn't found")
                    $retval = $false
                }
            }
            catch{
                wl($_.Exception.Message)
                $retval = $false
            }
            finally{
                if($retval){
                    wl("ValidatePackageState: $action SCCM_W10_APP_$($appname)_U exists")
                }
            }
        }
    }else{
        wl("Error: $appname application doesn't exist")
    }
    wl("ValidatePackageState: Completed")
}

[void]$FormSettings.ShowDialog()
$FormSettings.Visible = $false

[void]$Form.ShowDialog()
