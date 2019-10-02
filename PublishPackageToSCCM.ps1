<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    PublishPackageToSCCM2
.SYNOPSIS
    GUI based application. Publish package to SCCM, create AD Groups and distribute software.
.DESCRIPTION
    Use variables section to change main settings 
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1150,920'
$Form.text                       = "Publish Package to SCCM"
$Form.TopMost                    = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "PackageName"
$Label1.AutoSize                 = $true
$Label1.width                    = 254
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(19,24)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Publisher"
$Label2.AutoSize                 = $true
$Label2.width                    = 254
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(20,50)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "ApplicationName"
$Label3.AutoSize                 = $true
$Label3.width                    = 254
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(20,80)
$Label3.Font                     = 'Microsoft Sans Serif,10'

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Version"
$Label4.AutoSize                 = $true
$Label4.width                    = 254
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(20,110)
$Label4.Font                     = 'Microsoft Sans Serif,10'

$Groupbox1                       = New-Object system.Windows.Forms.Groupbox
$Groupbox1.height                = 45
$Groupbox1.width                 = 540
$Groupbox1.text                  = "Deployment Type"
$Groupbox1.location              = New-Object System.Drawing.Point(600,720)

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
$btnStart.location               = New-Object System.Drawing.Point(902,478)
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

$Label6                          = New-Object system.Windows.Forms.Label
$Label6.text                     = "Test Machines (;)"
$Label6.AutoSize                 = $true
$Label6.width                    = 25
$Label6.height                   = 10
$Label6.location                 = New-Object System.Drawing.Point(20,20)
$Label6.Font                     = 'Microsoft Sans Serif,10'

$tbTestMachine                   = New-Object system.Windows.Forms.TextBox
$tbTestMachine.multiline         = $false
$tbTestMachine.width             = 400
$tbTestMachine.height            = 20
$tbTestMachine.Anchor            = 'top,right,left'
$tbTestMachine.location          = New-Object System.Drawing.Point(135,20)
$tbTestMachine.Font              = 'Microsoft Sans Serif,10'

$Label7                          = New-Object system.Windows.Forms.Label
$Label7.text                     = "Test Users (;)"
$Label7.AutoSize                 = $true
$Label7.width                    = 25
$Label7.height                   = 10
$Label7.location                 = New-Object System.Drawing.Point(20,50)
$Label7.Font                     = 'Microsoft Sans Serif,10'

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
$lbLog.height                    = 210
$lbLog.Anchor                    = 'top,right,bottom,left'
$lbLog.location                  = New-Object System.Drawing.Point(10,10)

$ddApplications                  = New-Object system.Windows.Forms.ComboBox
$ddApplications.text             = "Click Get Applications for Selected Folder button to Populate this list"
$ddApplications.width            = 560
$ddApplications.height           = 20
$ddApplications.visible          = $true
$ddApplications.enabled          = $false
$ddApplications.Anchor           = 'right,bottom,left'
$ddApplications.location         = New-Object System.Drawing.Point(10,260)
$ddApplications.Font             = 'Microsoft Sans Serif,10'

$btnRefreshApplications          = New-Object system.Windows.Forms.Button
$btnRefreshApplications.text     = "Get Applications for Selected Folder"
$btnRefreshApplications.width    = 270
$btnRefreshApplications.height   = 30
$btnRefreshApplications.visible  = $true
$btnRefreshApplications.enabled  = $false
$btnRefreshApplications.Anchor   = 'bottom,left'
$btnRefreshApplications.location  = New-Object System.Drawing.Point(240,290)
$btnRefreshApplications.Font     = 'Microsoft Sans Serif,10'

$btnLoadApplication              = New-Object system.Windows.Forms.Button
$btnLoadApplication.text         = "Load Application"
$btnLoadApplication.width        = 150
$btnLoadApplication.height       = 30
$btnLoadApplication.visible      = $true
$btnLoadApplication.enabled      = $false
$btnLoadApplication.Anchor       = 'bottom,left'
$btnLoadApplication.location     = New-Object System.Drawing.Point(10,330)
$btnLoadApplication.Font         = 'Microsoft Sans Serif,10'

$ddApplicationFolders            = New-Object system.Windows.Forms.ComboBox
$ddApplicationFolders.text       = "Click Refresh Application Folders button to Populate this list"
$ddApplicationFolders.width      = 560
$ddApplicationFolders.height     = 20
$ddApplicationFolders.Anchor     = 'right,bottom,left'
$ddApplicationFolders.location   = New-Object System.Drawing.Point(10,230)
$ddApplicationFolders.Font       = 'Microsoft Sans Serif,10'

$btnRefreshApplicationFolders    = New-Object system.Windows.Forms.Button
$btnRefreshApplicationFolders.text  = "Refresh Application Folders"
$btnRefreshApplicationFolders.width  = 200
$btnRefreshApplicationFolders.height  = 30
$btnRefreshApplicationFolders.Anchor  = 'bottom,left'
$btnRefreshApplicationFolders.location  = New-Object System.Drawing.Point(10,290)
$btnRefreshApplicationFolders.Font  = 'Microsoft Sans Serif,10'

$btnRemoveApplication            = New-Object system.Windows.Forms.Button
$btnRemoveApplication.text       = "Remove Application and its Deployments"
$btnRemoveApplication.width      = 310
$btnRemoveApplication.height     = 30
$btnRemoveApplication.visible    = $true
$btnRemoveApplication.enabled    = $false
$btnRemoveApplication.Anchor     = 'bottom,left'
$btnRemoveApplication.location   = New-Object System.Drawing.Point(200,330)
$btnRemoveApplication.Font       = 'Microsoft Sans Serif,10'

$tbUnikeyRef                     = New-Object system.Windows.Forms.TextBox
$tbUnikeyRef.multiline           = $false
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

$Groupbox2                       = New-Object system.Windows.Forms.Groupbox
$Groupbox2.height                = 170
$Groupbox2.width                 = 540
$Groupbox2.text                  = "Package Properties"
$Groupbox2.location              = New-Object System.Drawing.Point(600,540)

$Groupbox3                       = New-Object system.Windows.Forms.Groupbox
$Groupbox3.height                = 80
$Groupbox3.width                 = 540
$Groupbox3.text                  = "Test"
$Groupbox3.location              = New-Object System.Drawing.Point(600,780)

$Groupbox4                       = New-Object system.Windows.Forms.Groupbox
$Groupbox4.height                = 370
$Groupbox4.width                 = 580
$Groupbox4.location              = New-Object System.Drawing.Point(10,540)

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

$cbDeployToUninstallDeviceCollectionsHiddenRequired   = New-Object system.Windows.Forms.CheckBox
$cbDeployToUninstallDeviceCollectionsHiddenRequired.text  = "Deploy To Uninstall Device Collections (Hidden, Required)"
$cbDeployToUninstallDeviceCollectionsHiddenRequired.AutoSize  = $true
$cbDeployToUninstallDeviceCollectionsHiddenRequired.width  = 95
$cbDeployToUninstallDeviceCollectionsHiddenRequired.height  = 20
$cbDeployToUninstallDeviceCollectionsHiddenRequired.location  = New-Object System.Drawing.Point(10,40)
$cbDeployToUninstallDeviceCollectionsHiddenRequired.Font  = 'Microsoft Sans Serif,10'

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
$cbMoveUninstallDeviceCollection.AutoSize  = $false
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

$Groupbox2.controls.AddRange(@($Label1,$Label2,$Label3,$Label4,$tbPackageName,$tbPublisher,$tbApplicationName,$tbVersion,$errPackageName,$errPublisher,$errApplicationName,$errVersion,$tbUnikeyRef,$lblUnikeyRef))
$Form.controls.AddRange(@($Groupbox1,$btnStart,$Groupbox2,$Groupbox3,$Groupbox4,$gbActions,$ddPresets))
$Groupbox1.controls.AddRange(@($rbMSI,$rbPSAppDeploy,$rbBAT))
$Groupbox3.controls.AddRange(@($Label6,$tbTestMachine,$Label7,$tbTestUser))
$Groupbox4.controls.AddRange(@($lbLog,$ddApplications,$btnRefreshApplications,$btnLoadApplication,$ddApplicationFolders,$btnRefreshApplicationFolders,$btnRemoveApplication))
$gbActions.controls.AddRange(@($gbCreatePackage,$gbCreateCollections,$gbDeployToCollections,$gbCreateADGroups,$gbCreateADGroupsQueries,$gbExtra,$gbTestMachines,$gbTestUsers,$gbRemoveCollections,$gbRemoveADGroups))
$gbCreatePackage.controls.AddRange(@($cbCreatePackage,$cbMovePackage,$cbCreateDeploymentType,$cbDistributeContent))
$gbCreateCollections.controls.AddRange(@($cbCreateInstallDeviceCollection,$cbCreateUninstallDeviceCollection,$cbCreateAvailableUserCollection,$cbMoveInstallDeviceCollection,$cbMoveUninstallDeviceCollection,$cbMoveAvailableUserCollection))
$gbDeployToCollections.controls.AddRange(@($cbDeployToInstallDeviceCollectionHiddenRequired,$cbDeployToUninstallDeviceCollectionsHiddenRequired,$cbDeployToUserCollectionAvailable))
$gbCreateADGroups.controls.AddRange(@($cbCreateInstallADGroup,$cbCreateUninstallADGroup,$cbCreateUserAvailableADGroup))
$gbCreateADGroupsQueries.controls.AddRange(@($cbAddADGroupQueryToInstallDeviceCollection,$cbAddADGroupQueryToUninstallDeviceCollection,$cbAddADGroupQueryToAvailableUserCollection))
$gbExtra.controls.AddRange(@($cbUpdateContentOnDPs))
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

$DetectionKeyLocation = "SOFTWARE\CompanyName\Installed" # Standard Audit key is located HKLM\SOFTWARE\CompanyName\Installed\$PackageName, key name Installed, value Date and time of installation




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
        try{
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
        }
        catch{
            wl("Failed to import SCCM module. Please make sure that you have SCCM Console installed")
            wl($_.Exception.Message)
        }
    }
    wl("Connected to SCCM")
    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ActiveDirectory) -eq $null) {
        try{
            Import-Module ActiveDirectory
        }
        catch{
            wl("Failed to import ActiveDirectory module. Please make sure that you have RSAT installed")
            wl($_.Exception.Message)
        }
    }
    wl("Connected to AD")
    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        try{
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer
        }
        catch{
            wl($_.Exception.Message)
        }
    }
    # Check Package Repository can be connected
    if(![System.IO.Directory]::Exists($PackageRepository)){
        wl("Can't connect to Package repository $PackageRepository")
    }
    # Set the current location to be the site code.
    try{
        Set-Location "$($SiteCode):\"
    }
    catch{
        wl($_.Exception.Message)
    }
}

function FormShown { ConnectToSCCMAndAD }


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
    $cbDeployToUninstallDeviceCollectionsHiddenRequired.checked = $false
    $cbDeployToUserCollectionAvailable.checked = $false
    $cbCreateInstallADGroup.checked = $false
    $cbCreateUninstallADGroup.checked = $false
    $cbCreateUserAvailableADGroup.checked = $false
    $cbAddADGroupQueryToInstallDeviceCollection.checked = $false
    $cbAddADGroupQueryToUninstallDeviceCollection.checked = $false
    $cbAddADGroupQueryToAvailableUserCollection.checked = $false
    $cbUpdateContentOnDPs.checked = $false
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
    $cbDeployToUninstallDeviceCollectionsHiddenRequired.checked = $true
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
    $MSIName = ""
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
    }
    #Move Package
    if($cbMovePackage.checked -And $returnerror -eq ""){
	    if($NewPackageLocation -ne ""){
            wl("Move Package: Move-CMObject -InputObject `$(Get-CMApplication -Name ""$PackageName"") -FolderPath ""$($SiteCode):\Application\$NewPackageLocation""")
            try{
    			Move-CMObject -InputObject $(Get-CMApplication -Name "$PackageName") -FolderPath "$($SiteCode):\Application\$NewPackageLocation"
    			wl("Move Package: moved application to $NewPackageLocation")
    		}
    		catch{
    		    $returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
	    }
    }
    #Create Deployment Type
    if($cbCreateDeploymentType.checked -And $returnerror -eq ""){
        if($rbMSI.checked){
            $MSIName = GetMSIFileName($defaultPackageLocation)
            if($MSIName -eq ""){
                $returnerror += ShowMessageBoxWithError("No MSI files were found in $defaultPackageLocation folder")
            }else{
                wl("Create Deployment Type: found msi file $MSIName")
                wl("Create Deployment Type: getting ProductCode property for $defaultPackageLocation\$($MSIName)")
                $tmp = GetProductCode("$defaultPackageLocation\$($MSIName)")
                $ProductCode = ($tmp|Out-String).Trim()
                if($ProductCode -like "*error*"){
                    $returnerror += $ProductCode
                }else{
                    wl("Create Deployment Type: ProductCode=$ProductCode")
                    $guidProductCode = [GUID]$ProductCode
                	$InstallCommand = "msiexec.exe /i ""$($MSIName)"""
                	If([System.IO.File]::Exists("$defaultPackageLocation\$($PackageName).mst")){$InstallCommand += " TRANSFORMS=$($PackageName).mst"}
                	$InstallCommand += " /qn /l* C:\Windows\Logs\$($PackageName)_I.log"
                	wl("Create Deployment Type: InstallCommand=$InstallCommand")
                	$UninstallCommand = "msiexec.exe /x $ProductCode /qn /l* C:\Windows\Logs\$($PackageName)_U.log"
                	wl("Create Deployment Type: UninstallCommand=$UninstallCommand")
                	wl("Create Deployment Type: Add-CMMsiDeploymentType -DeploymentTypeName `"$PackageName`" -InstallCommand `"$InstallCommand`" -ApplicationName `"$PackageName`" -ProductCode $ProductCode -ContentLocation `"$defaultPackageLocation`" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand `"$UninstallCommand`" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment `"$description`" -Force")
                    try{
                    	Add-CMMsiDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "$InstallCommand" -ApplicationName "$PackageName" -ProductCode $ProductCode -ContentLocation "$defaultPackageLocation\$($MSIName)" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "$UninstallCommand" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description" -Force
                    	wl("Create Deployment Type: created MSI deployment type")
                    }catch{
                    	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
                    }
                }
            }
        }
        if($rbPSAppDeploy.checked){
            wl("Create Deployment Type: Add-CMScriptDeploymentType -DeploymentTypeName ""$PackageName"" -InstallCommand ""Deploy-Application.exe -DeploymentType """"Install"""" -DeployMode """"Silent"""""" -ApplicationName ""$PackageName"" -AddDetectionClause `$(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName ""$DetectionKeyLocation\$PackageName"" -PropertyType String -ValueName ""Installed"" -Is64Bit -Existence) -ContentLocation ""$defaultPackageLocation"" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand ""Deploy-Application.exe -DeploymentType """"Uninstall"""" -DeployMode """"Silent"""""" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment ""$description""")
            try{
            	Add-CMScriptDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "Deploy-Application.exe -DeploymentType ""Install"" -DeployMode ""Silent""" -ApplicationName "$PackageName" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName "$DetectionKeyLocation\$PackageName" -PropertyType String -ValueName "Installed" -Is64Bit -Existence) -ContentLocation "$defaultPackageLocation" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "Deploy-Application.exe -DeploymentType ""Uninstall"" -DeployMode ""Silent""" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description"
            	wl("Create Deployment Type: created PSAppdeploy scripted deployment type")
            }catch{
            	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
            }
        }
        if($rbBAT.checked){
            wl("Create Deployment Type: Add-CMScriptDeploymentType -DeploymentTypeName ""$PackageName"" -InstallCommand ""install.bat"" -ApplicationName ""$PackageName"" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName ""$DetectionKeyLocation\$PackageName"" -PropertyType String -ValueName ""Installed"" -Is64Bit -Existence) -ContentLocation ""$defaultPackageLocation"" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand ""uninstall.bat"" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment ""$description""")
            try{
            	Add-CMScriptDeploymentType -DeploymentTypeName "$PackageName" -InstallCommand "install.bat" -ApplicationName "$PackageName" -AddDetectionClause $(New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName "$DetectionKeyLocation\$PackageName" -PropertyType String -ValueName "Installed" -Is64Bit -Existence) -ContentLocation "$defaultPackageLocation" -LogonRequirementType WhetherOrNotUserLoggedOn -UninstallCommand "uninstall.bat" -UserInteractionMode Hidden -InstallationBehaviorType InstallForSystem -Comment "$description"
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
    #Create Install Device Collection
    if($cbCreateInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Create Install Device Collection: New-CMDeviceCollection -Name ""$InstallCollectionName"" -LimitingCollectionName ""$LimitingDeviceCollectionName"" -RefreshType Continuous -Comment ""$description""")
        try{
        	New-CMDeviceCollection -Name "$InstallCollectionName" -LimitingCollectionName "$LimitingDeviceCollectionName" -RefreshType Continuous -Comment "$description"
        	wl("Create Install Device Collections: created Install Device Collection $InstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Move Install Device Collection
    if($cbMoveInstallDeviceCollection.checked -And $returnerror -eq "" -And $InstallDeviceCollectionLocation -ne ""){
        wl("Move Install Device Collection: Move-CMObject -InputObject `$(Get-CMDeviceCollection -Name ""$InstallCollectionName"") -FolderPath ""$($SiteCode):\DeviceCollection\$InstallDeviceCollectionLocation""")
        try{
        	Move-CMObject -InputObject $(Get-CMDeviceCollection -Name "$InstallCollectionName") -FolderPath "$($SiteCode):\DeviceCollection\$InstallDeviceCollectionLocation"
        	wl("Move Install Device Collection: moved it to $InstallDeviceCollectionLocation")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Create Uninstall Device Collection
    if($cbCreateUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Create Uninstall Device Collection: New-CMDeviceCollection -Name ""$UninstallCollectionName"" -LimitingCollectionName ""$LimitingDeviceCollectionName"" -RefreshType Continuous -Comment ""$description""")
        try{
        	New-CMDeviceCollection -Name "$UninstallCollectionName" -LimitingCollectionName "$LimitingDeviceCollectionName" -RefreshType Continuous -Comment "$description"
        	wl("Create Uninstall Device Collection: created Uninstall Device Collection $UninstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Move Uninstall Device Collection
    if($cbMoveUninstallDeviceCollection.checked -And $returnerror -eq "" -And $UninstallDeviceCollectionLocation -ne ""){
        wl("Move Uninstall Device Collections: Move-CMObject -InputObject `$(Get-CMDeviceCollection -Name ""$UninstallCollectionName"") -FolderPath ""$($SiteCode):\DeviceCollection\$UninstallDeviceCollectionLocation""")
        try{
        	Move-CMObject -InputObject $(Get-CMDeviceCollection -Name "$UninstallCollectionName") -FolderPath "$($SiteCode):\DeviceCollection\$UninstallDeviceCollectionLocation"
        	wl("Move Uninstall Device Collections: moved it to $UninstallDeviceCollectionLocation")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Create Available User Collection
    if($cbCreateAvailableUserCollection.checked -And $returnerror -eq ""){
        wl("Create Available User Collection: New-CMUserCollection -Name ""$UserCollectionName"" -LimitingCollectionName ""$LimitingUserCollectionName"" -RefreshType Continuous")
        try{
        	New-CMUserCollection -Name "$UserCollectionName" -LimitingCollectionName "$LimitingUserCollectionName" -RefreshType Continuous
        	wl("Create Available User Collection: created Install User Collection $UserCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Move Available User Collection
    if($cbMoveAvailableUserCollection.checked -And $returnerror -eq "" -And $UserCollectionLocation -ne ""){
        wl("Move Available User Collection: Move-CMObject -InputObject $(Get-CMUserCollection -Name ""$UserCollectionName"") -FolderPath ""$($SiteCode):\UserCollection\$UserCollectionLocation""")
        try{
        	Move-CMObject -InputObject $(Get-CMUserCollection -Name "$UserCollectionName") -FolderPath "$($SiteCode):\UserCollection\$UserCollectionLocation"
    	    wl("Move Available User Collection: moved it to $UserCollectionLocation")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Deploy To Install Device Collection (Required, Hidden)
    if($cbDeployToInstallDeviceCollectionHiddenRequired.checked -And $returnerror -eq ""){
        wl("Deploy To Install Device Collection (Required, Hidden): New-CMApplicationDeployment -Name ""$PackageName"" -DeployAction Install -DeployPurpose Required -UserNotification HideAll -CollectionName ""$InstallCollectionName"" -Comment ""$description""")
        try{
        	New-CMApplicationDeployment -Name "$PackageName" -DeployAction Install -DeployPurpose Required -UserNotification HideAll -CollectionName "$InstallCollectionName" -Comment "$description"
        	wl("Deploy To Install Device Collection (Required, Hidden): assigned application for deployment to $InstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Deploy To Uninstall Device Collection (Required, Hidden)
    if($cbDeployToInstallDeviceCollectionHiddenRequired.checked -And $returnerror -eq ""){
        wl("Deploy To Uninstall Device Collection (Required, Hidden): New-CMApplicationDeployment -Name ""$PackageName"" -DeployAction Uninstall -DeployPurpose Required -UserNotification HideAll -CollectionName ""$UninstallCollectionName"" -Comment ""$description""")
        try{
        	New-CMApplicationDeployment -Name "$PackageName" -DeployAction Uninstall -DeployPurpose Required -UserNotification HideAll -CollectionName "$UninstallCollectionName" -Comment "$description"
        	wl("Deploy To Uninstall Device Collection (Required, Hidden): assigned application for deployment to $UninstallCollectionName")
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
    #Create Install AD Group (_I)
    if($cbCreateInstallADGroup.checked -And $returnerror -eq ""){
        wl("Create Install AD Group (_I): New-ADGroup ""$ADGroupNamePrefix$($InstallCollectionName)"" -Path ""$ADOUPath"" -GroupCategory Security -GroupScope Global -Description ""$localizedName"" -OtherAttributes @{info=""This group deploys $($tbUnikeyRef.text) """"$PackageName"""" to any Windows 10 devices specified - Please do not add users to the group. Devices should only be added with the necessary approval. If in doubt please liaise with the application owner.""}")
		try{
			New-ADGroup "$ADGroupNamePrefix$($InstallCollectionName)" -Path "$ADOUPath" -GroupCategory Security -GroupScope Global -Description "$localizedName" -OtherAttributes @{info="This group deploys $($tbUnikeyRef.text) `"$PackageName`" to any Windows 10 devices specified - Please do not add users to the group. Devices should only be added with the necessary approval. If in doubt please liaise with the application owner."}
        	wl("Create Install AD Group (_I): created AD group $ADGroupNamePrefix$($InstallCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Create Uninstall AD Group (_U)
    if($cbCreateUninstallADGroup.checked -And $returnerror -eq ""){
        wl("Create Uninstall AD Group (_U): New-ADGroup ""$ADGroupNamePrefix$($UninstallCollectionName)"" -Path ""$ADOUPath"" -GroupCategory Security -GroupScope Global -Description ""$localizedName""")
		try{
			New-ADGroup "$ADGroupNamePrefix$($UninstallCollectionName)" -Path "$ADOUPath" -GroupCategory Security -GroupScope Global -Description "$localizedName" 
        	wl("Create Uninstall AD Group (_U): created AD group $ADGroupNamePrefix$($UninstallCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Create User Available AD Group (_A)
    if($cbCreateUserAvailableADGroup.checked -And $returnerror -eq ""){
        wl("Create User Available AD Group (_A): New-ADGroup ""$ADGroupNamePrefix$($UserCollectionName)"" -Path ""$ADOUPath"" -GroupCategory Security -GroupScope Global -Description ""$localizedName"" -OtherAttributes @{info=""This group makes available $($tbUnikeyRef.text) """"$PackageName"""" for users specified - Please do not add computer to this group. Users should only be added with the necessary approval. If in doubt please liaise with the application owner.""}")
		try{
			New-ADGroup "$ADGroupNamePrefix$($UserCollectionName)" -Path "$ADOUPath" -GroupCategory Security -GroupScope Global -Description "$localizedName" -OtherAttributes @{info="This group makes available $($tbUnikeyRef.text) `"$PackageName`" for users specified - Please do not add computer to this group. Users should only be added with the necessary approval. If in doubt please liaise with the application owner."}
        	wl("Create User Available AD Group (_A): created AD group $ADGroupNamePrefix$($UserCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Add AD Groups Query to Install Device Collection
    if($cbAddADGroupQueryToInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Add AD Groups Query to Install Device Collection: Add-CMDeviceCollectionQueryMembershipRule -CollectionName ""$InstallCollectionName"" -RuleName ""$ADGroupNamePrefix$($InstallCollectionName)"" -QueryExpression ""$ADGroupQuery$ADGroupNamePrefix$($InstallCollectionName)'""")
		try{
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$InstallCollectionName" -RuleName "$ADGroupNamePrefix$($InstallCollectionName)" -QueryExpression "$ADGroupQuery$ADGroupNamePrefix$($InstallCollectionName)'"
        	wl("Add AD Groups Query to Install Device Collection: created query for AD group $ADGroupNamePrefix$($InstallCollectionName) in $InstallCollectionName")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Add AD Groups Query to Uninstall Device Collection
    if($cbAddADGroupQueryToUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Add AD Groups Query to Uninstall Device Collection: Add-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UninstallCollectionName"" -RuleName ""$ADGroupNamePrefix$($UninstallCollectionName)"" -QueryExpression ""$ADGroupQuery$ADGroupNamePrefix$($UninstallCollectionName)'""")
		try{
			Add-CMDeviceCollectionQueryMembershipRule -CollectionName "$UninstallCollectionName" -RuleName "$ADGroupNamePrefix$($UninstallCollectionName)" -QueryExpression "$ADGroupQuery$ADGroupNamePrefix$($UninstallCollectionName)'"
        	wl("Add AD Groups Query to Uninstall Device Collection: created query for AD group $ADGroupNamePrefix$($UninstallCollectionName) in $UninstallCollectionName")
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
    #Update Content On Distribution Points
    if($cbUpdateContentOnDPs.checked -And $returnerror -eq ""){
        wl("Update Content On DPs: Enumirating Deployment Types for $PackageName application")
		try{
            foreach($dt in $(Get-CMDeploymentType -ApplicationName "$PackageName")){
               wl("Update Content On DPs: Update-CMDistributionPoint -ApplicationName `"$PackageName`" -DeploymentTypeName `"$($dt.LocalizedDisplayName)`"")
               Update-CMDistributionPoint -ApplicationName "$PackageName" -DeploymentTypeName "$($dt.LocalizedDisplayName)"
        	   wl("Update Content On DPs: updated distributed content for $($dt.LocalizedDisplayName)")
            }            
			
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Add Test Machines To Install Device Collection
    if($cbAddTestMachinesToInstallDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("Add Test Machines To Install Device Collection: Add-CMDeviceCollectionDirectMembershipRule -CollectionName ""$InstallCollectionName"" -Resource `$(Get-CMDevice -Name $tm)")
    		try{
    			Add-CMDeviceCollectionDirectMembershipRule -CollectionName "$InstallCollectionName" -Resource $(Get-CMDevice -Name $tm)
        	    wl("Add Test Machines To Install Device Collection: added $tm to $InstallCollectionName")
    		}catch{
    			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
        }
    }
    #Remove Test Machines From Install Device Collection
    if($cbRemoveTestMachinesFromInstallDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("Remove Test Machines From Install Device Collection: Remove-CMDeviceCollectionDirectMembershipRule -CollectionName ""$InstallCollectionName"" -ResourceName $tm -Force")
    		try{
    			Remove-CMDeviceCollectionDirectMembershipRule -CollectionName "$InstallCollectionName" -ResourceName $tm -Force
        	    wl("Remove Test Machines From Install Device Collection: removed $tm from $InstallCollectionName")
    		}catch{
    			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
        }
    }
    #Add Test Machines To Uninstall Device Collection
    if($cbAddTestMachinesToUninstallDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("Add Test Machines To Uninstall Device Collection: Add-CMDeviceCollectionDirectMembershipRule -CollectionName ""$UninstallCollectionName"" -Resource `$(Get-CMDevice -Name $tm)")
    		try{
    			Add-CMDeviceCollectionDirectMembershipRule -CollectionName "$UninstallCollectionName" -Resource $(Get-CMDevice -Name $tm)
        	    wl("Add Test Machines To Uninstall Device Collection: added $tm to $UninstallCollectionName")
    		}catch{
    			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
        }
    }
    #Remove Test Machines From Uninstall Device Collection
    if($cbRemoveTestMachinesFromUninstallDeviceCollection.checked -And $returnerror -eq ""){
        foreach($tm in $testmachine){
            wl("Remove Test Machines From Uninstall Device Collection: Remove-CMDeviceCollectionDirectMembershipRule -CollectionName ""$UninstallCollectionName"" -ResourceName $tm -Force")
    		try{
    			Remove-CMDeviceCollectionDirectMembershipRule -CollectionName "$UninstallCollectionName" -ResourceName $tm -Force
        	    wl("Remove Test Machines From Uninstall Device Collection: removed $tm from $UninstallCollectionName")
    		}catch{
    			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
    		}
        }
    }
    #Add Test Users To Available User Collection
    if($cbAddTestUsersToAvailableUserCollection.checked -And $returnerror -eq ""){
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
    if($cbRemoveTestUsersFromAvailableUserCollection.checked -And $returnerror -eq ""){
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
    #Remove Install Device Collection
    if($cbRemoveInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Remove Install Device Collection: Remove-CMDeviceCollection -Name ""$InstallCollectionName"" -Force")
        try{
        	Remove-CMDeviceCollection -Name "$InstallCollectionName" -Force
        	wl("Remove Install Device Collections: Removed Install Device Collection $InstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Remove Uninstall Device Collection
    if($cbRemoveUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Remove Uninstall Device Collection: Remove-CMDeviceCollection -Name ""$UninstallCollectionName"" -Force")
        try{
        	Remove-CMDeviceCollection -Name "$UninstallCollectionName" -Force
        	wl("Remove Uninstall Device Collection: Removed Uninstall Device Collection $UninstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Remove Available User Collection
    if($cbRemoveAvailableUserCollection.checked -And $returnerror -eq ""){
        wl("Remove Available User Collection: Remove-CMUserCollection -Name ""$UserCollectionName"" -Force")
        try{
        	Remove-CMUserCollection -Name "$UserCollectionName" -Force
        	wl("Remove Available User Collection: Removed Install User Collection $UserCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Remove Deployment To Install Device Collection
    if($cbRemoveDeploymentToInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Remove Deployment ToInstall Device Collection: Remove-CMApplicationDeployment -Name ""$PackageName"" -CollectionName ""$InstallCollectionName"" -Force")
        try{
        	Remove-CMApplicationDeployment -Name "$PackageName" -CollectionName "$InstallCollectionName" -Force
        	wl("Remove Deployment To Install Device Collections: Removed Install Device Collection $InstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Remove Deployment To Uninstall Device Collection
    if($cbRemoveDeploymentToUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Remove Deployment To Uninstall Device Collection: Remove-CMApplicationDeployment -Name ""$PackageName"" -CollectionName ""$UninstallCollectionName"" -Force")
        try{
        	Remove-CMApplicationDeployment -Name "$PackageName" -CollectionName "$UninstallCollectionName" -Force
        	wl("Remove Deployment To Uninstall Device Collection: Removed Uninstall Device Collection $UninstallCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Remove Deployment To Available User Collection
    if($cbRemoveDeploymentToAvailableUserCollection.checked -And $returnerror -eq ""){
        wl("Remove Deployment To Available User Collection: Remove-CMApplicationDeployment -Name ""$PackageName"" -CollectionName ""$UserCollectionName"" -Force")
        try{
        	Remove-CMApplicationDeployment -Name "$PackageName" -CollectionName "$UserCollectionName" -Force
        	wl("Remove Deployment To Available User Collection: Removed Install User Collection $UserCollectionName")
        }catch{
        	$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
        }
    }
    #Remove Install AD Group (_I)
    if($cbRemoveInstallADGroup.checked -And $returnerror -eq ""){
        wl("Remove Install AD Group (_I): Remove-ADGroup ""$ADGroupNamePrefix$($InstallCollectionName)""")
		try{
			Remove-ADGroup "$ADGroupNamePrefix$($InstallCollectionName)"
        	wl("Remove Install AD Group (_I): Removed AD group $ADGroupNamePrefix$($InstallCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Remove Uninstall AD Group (_U)
    if($cbRemoveUninstallADGroup.checked -And $returnerror -eq ""){
        wl("Remove Uninstall AD Group (_U): Remove-ADGroup ""$ADGroupNamePrefix$($UninstallCollectionName)""")
		try{
			Remove-ADGroup "$ADGroupNamePrefix$($UninstallCollectionName)"
        	wl("Remove Uninstall AD Group (_U): Removed AD group $ADGroupNamePrefix$($UninstallCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Remove User Available AD Group (_A)
    if($cbRemoveUserAvailableADGroup.checked -And $returnerror -eq ""){
        wl("Remove User Available AD Group (_A): Remove-ADGroup ""$ADGroupNamePrefix$($UserCollectionName)""")
		try{
			Remove-ADGroup "$ADGroupNamePrefix$($UserCollectionName)"
        	wl("Remove User Available AD Group (_A): Removed AD group $ADGroupNamePrefix$($UserCollectionName)")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Remove AD Groups Query to Install Device Collection
    if($cbRemoveADGroupQueryToInstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Remove AD Groups Query to Install Device Collection: Remove-CMDeviceCollectionQueryMembershipRule -CollectionName ""$InstallCollectionName"" -RuleName ""$ADGroupNamePrefix$($InstallCollectionName)"" -Force")
		try{
			Remove-CMDeviceCollectionQueryMembershipRule -CollectionName "$InstallCollectionName" -RuleName "$ADGroupNamePrefix$($InstallCollectionName)" -Force
        	wl("Remove AD Groups Query to Install Device Collection: removed query for AD group $ADGroupNamePrefix$($InstallCollectionName) in $InstallCollectionName")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Remove AD Groups Query to Uninstall Device Collection
    if($cbRemoveADGroupQueryToUninstallDeviceCollection.checked -And $returnerror -eq ""){
        wl("Remove AD Groups Query to Uninstall Device Collection: Remove-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UninstallCollectionName"" -RuleName ""$ADGroupNamePrefix$($UninstallCollectionName)"" -Force")
		try{
			Remove-CMDeviceCollectionQueryMembershipRule -CollectionName "$UninstallCollectionName" -RuleName "$ADGroupNamePrefix$($UninstallCollectionName)" -Force
        	wl("Remove AD Groups Query to Uninstall Device Collection: removed query for AD group $ADGroupNamePrefix$($UninstallCollectionName) in $UninstallCollectionName")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
		}
    }
    #Remove AD Group Query To User Collection
    if($cbRemoveADGroupQueryToUserCollection.checked -And $returnerror -eq ""){
        wl("Remove AD Group Query To User Collection: Remove-CMDeviceCollectionQueryMembershipRule -CollectionName ""$UserCollectionName"" -RuleName ""$ADGroupNamePrefix$($UserCollectionName)"" -Force")
		try{
			Remove-CMDeviceCollectionQueryMembershipRule -CollectionName "$UserCollectionName" -RuleName "$ADGroupNamePrefix$($UserCollectionName)" -Force
        	wl("Remove AD Group Query To User Collection: removed query for AD group $ADGroupNamePrefix$($UserCollectionName) in $UserCollectionName")
		}catch{
			$returnerror += ShowMessageBoxWithError("Error: "+ $_.Exception.Message)
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
    $Instancekeys = (Get-WmiObject -Namespace "ROOT\SMS\Site_$Sitecode" -ComputerName $SiteServer `
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

[void]$Form.ShowDialog()
