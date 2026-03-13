Imports System.Diagnostics.Contracts
Imports System.IO
Imports System.Net
Imports System.Net.Mime.MediaTypeNames
Imports System.Reflection.Metadata
Imports System.Resources
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices.JavaScript.JSType
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.VisualBasic.Logging

Public Class SysTray
   Inherits ApplicationContext

   Private tray As NotifyIcon
   Private mnuMain As ContextMenuStrip

   Public Sub New()
	  ' Build menu
	  mnuMain = New ContextMenuStrip()

	  ' Shell Folders -----------------------------------------------------------------------------

	  Dim mnuShellFolders As New ToolStripMenuItem("Shell Folders")

	  Dim dicShellFolderscategorized As New Dictionary(Of String, Dictionary(Of String, String)) From {
		 {"Current User", New Dictionary(Of String, String) From {
			{"Profile", "Profile"},
			{"3D Objects", "3D Objects"},
			{"Contacts", "Contacts"},
			{"Desktop", "Desktop"},
			{"Downloads", "Downloads"},
			{"Favorites", "Favorites"},
			{"Links", "Links"},
			{"Music", "My Music"},
			{"Personal", "Personal"},
			{"Pictures", "My Pictures"},
			{"Camera Roll", "Camera Roll"},
			{"Screenshots", "Screenshots"},
			{"Videos", "My Video"},
			{"Recycle Bin Folder", "RecycleBinFolder"},
			{"Ringtones", "Ringtones"},
			{"Saved Games", "SavedGames"},
			{"Searches", "Searches"},
			{"SendTo", "SendTo"},
			{"Libraries", "Libraries"},
			{"Documents Library", "DocumentsLibrary"},
			{"Music Library", "MusicLibrary"},
			{"Pictures Library", "PicturesLibrary"},
			{"Videos Library", "VideosLibrary"},
			{"Users Libraries Folder", "UsersLibrariesFolder"},
			{"User Pinned", "User Pinned"},
			{"User Profiles", "UserProfiles"},
			{"User ProgramFiles", "UserProgramFiles"},
			{"User ProgramFiles Common", "UserProgramFilesCommon"},
			{"User Tiles", "UserTiles"},
			{"Users Files Folder", "UsersFilesFolder"}
		 }},
		 {"Current User AppData", New Dictionary(Of String, String) From {
			{"AppData", "AppData"},
			{"Local AppData", "Local AppData"},
			{"Local AppData Low", "LocalAppDataLow"},
			{"Account Pictures", "AccountPictures"},
			{"Application Shortcuts", "Application Shortcuts"},
			{"Cache", "Cache"},
			{"CD Burning", "CD Burning"},
			{"Cookies", "Cookies"},
			{"Cookies\Low", "Cookies\Low"},
			{"Credential Manager", "CredentialManager"},
			{"Cryptokeys", "Cryptokeys"},
			{"Dp API Keys", "DpAPIKeys"},
			{"History", "History"},
			{"GameTasks", "GameTasks"},
			{"Implicit App Shortcuts", "ImplicitAppShortcuts"},
			{"PrintHood", "PrintHood"},
			{"Recent", "Recent"},
			{"Start Menu", "Start Menu"},
			{"Startup", "Startup"},
			{"System Certificates", "SystemCertificates"},
			{"Templates", "Templates"}
		 }},
		{"Public", New Dictionary(Of String, String) From {
			{"Common Desktop", "Common Desktop"},
			{"Common Documents", "Common Documents"},
			{"Common Downloads", "CommonDownloads"},
			{"Common Music", "CommonMusic"},
			{"Common Pictures", "CommonPictures"},
			{"Common Video", "CommonVideo"},
			{"Common Start Menu", "Common Start Menu"},
			{"Common Startup", "Common Startup"},
			{"Public", "Public"},
			{"Public Account Pictures", "PublicAccountPictures"},
			{"Public Libraries", "PublicLibraries"},
			{"Public User Tiles", "PublicUserTiles"}
		 }},
		{"Control Panel", New Dictionary(Of String, String) From {
		   {"Control Panel Folder", "ControlPanelFolder"},
		   {"Add New Programs Folder", "AddNewProgramsFolder"},
		   {"Administrative Tools", "Administrative Tools"},
		   {"Apps Folder", "AppsFolder"},
		   {"Change Remove Programs Folder", "ChangeRemoveProgramsFolder"},
		   {"Connections Folder", "ConnectionsFolder"},
		   {"NetHood", "NetHood"},
		   {"Printers Folder", "PrintersFolder"},
		   {"SyncCenter Folder", "SyncCenterFolder"},
		   {"SyncCenter Results Folder", "SyncResultsFolder"},
		   {"SyncCenter Setup Folder", "SyncSetupFolder"},
		   {"SyncCenter Conflict Folder", "ConflictFolder"},
		   {"SyncCenter CSC (Client Side Caching) Folder", "CSCFolder"},
		   {"Windows", "Windows"}
		}},
		{"Settings", New Dictionary(Of String, String) From {
		   {"App Updates Folder", "AppUpdatesFolder"}
		}},
		{"Windows", New Dictionary(Of String, String) From {
		   {"Fonts", "Fonts"},
		   {"My Computer Folder", "MyComputerFolder"},
		   {"Network Places Folder", "NetworkPlacesFolder"},
		   {"System (System32)", "System"},
		   {"System X86 (SysWOW64)", "SystemX86"},
		   {"Windows", "Windows"}
		}},
		{"Program Files", New Dictionary(Of String, String) From {
		   {"ProgramFiles", "ProgramFiles"},
		   {"ProgramFiles Common", "ProgramFilesCommon"},
		   {"ProgramFiles CommonX64", "ProgramFilesCommonX64"},
		   {"ProgramFiles CommonX86", "ProgramFilesCommonX86"},
		   {"ProgramFiles X64", "ProgramFilesX64"},
		   {"ProgramFiles X86", "ProgramFilesX86"},
		   {"Programs", "Programs"}
		}},
		{"ProgramData", New Dictionary(Of String, String) From {
		   {"Common AppData", "Common AppData"},
		   {"Common Administrative Tools", "Common Administrative Tools"},
		   {"Common Programs", "Common Programs"},
		   {"Common Ringtones", "CommonRingtones"},
		   {"Common Templates", "Common Templates"},
		   {"Device Metadata Store", "Device Metadata Store"},
		   {"Public Game Tasks", "PublicGameTasks"}
		}},
		{"OneDrive", New Dictionary(Of String, String) From {
		   {"OneDrive", "OneDrive"},
		   {"OneDrive CameraRoll", "OneDriveCameraRoll"},
		   {"OneDrive Documents", "OneDriveDocuments"},
		   {"OneDrive Music", "OneDriveMusic"},
		   {"OneDrive Pictures", "OneDrivePictures"},
		   {"SkyDrive CameraRoll ", "SkyDriveCameraRoll"},
		   {"SkyDrive Documents ", "SkyDriveDocuments"},
		   {"SkyDrive Music ", "SkyDriveMusic"},
		   {"SkyDrive Pictures ", "SkyDrivePictures"}
		}}
	  }

	  Dim dicShellFoldersUncategorized As New Dictionary(Of String, String) From {
		 {"Default Gadgets (Win Vista/7/8)", "Default Gadgets"}, 'Windows Vista, Windows 7, and Windows 8		
		 {"Gadgets (Win 7)", "Gadgets"}, 'Windows 7 only
		 {"Games", "Games"},
		 {"HomeGroup Current User Folder", "HomeGroupCurrentUserFolder"},
		 {"HomeGroup Folder", "HomeGroupFolder"},
		 {"Internet Folder", "InternetFolder"}, ' deprecated		 
		 {"Localized Resources Dir", "LocalizedResourcesDir"},
		 {"MAPI Folder ", "MAPIFolder"},
		 {"OEM Links", "OEM Links"},
		 {"Original Images", "Original Images"},
		 {"Photo Albums", "PhotoAlbums"},
		 {"Playlists", "Playlists"},
		 {"Quick Launch", "Quick Launch"},
		 {"Recorded TV Library", "RecordedTVLibrary"},
		 {"Resource Dir", "ResourceDir"},
		 {"Retail Demo", "Retail Demo"},
		 {"Roamed Tile Images", "Roamed Tile Images"},
		 {"Roaming Tiles", "Roaming Tiles"},
		 {"Sample Music", "SampleMusic"},
		 {"Sample Pictures", "SamplePictures"},
		 {"Sample Videos", "SampleVideos"},
		 {"Search History Folder", "SearchHistoryFolder"},
		 {"Search Home Folder", "SearchHomeFolder"},
		 {"Search Templates Folder", "SearchTemplatesFolder"},
		 {"Start Menu All Programs", "StartMenuAllPrograms"},
		 {"This PC Desktop Folder", "ThisPCDesktopFolder"}
	  }

	  For Each category In dicShellFolderscategorized.Keys
		 Dim categoryItem As New ToolStripMenuItem(category)
		 For Each entry In dicShellFolderscategorized(category)
			Dim displayName As String = entry.Key
			Dim shellName As String = entry.Value
			Dim item As New ToolStripMenuItem(displayName)
			Dim localShell = shellName
			AddHandler item.Click, Sub() LaunchShellCommands(localShell)
			categoryItem.DropDownItems.Add(item)
		 Next
		 mnuShellFolders.DropDownItems.Add(categoryItem)
	  Next

	  For Each entry In dicShellFoldersUncategorized
		 Dim item As New ToolStripMenuItem(entry.Key)
		 Dim localShell = entry.Value
		 AddHandler item.Click, Sub() LaunchShellCommands(localShell)
		 mnuShellFolders.DropDownItems.Add(item)
	  Next

	  ' Build main menu
	  mnuMain.Items.Add(mnuShellFolders)
	  mnuMain.Items.Add(New ToolStripSeparator())
	  mnuMain.Items.Add("&About", Nothing, AddressOf AppAbout)
	  mnuMain.Items.Add("Rebuild missing folders", Nothing, AddressOf AppRebuild)
	  mnuMain.Items.Add("E&xit", Nothing, AddressOf AppExit)

	  ' Create tray icon
	  tray = New NotifyIcon()
	  tray.Icon = My.Resources.Resources.terminal
	  tray.Visible = True
	  tray.ContextMenuStrip = mnuMain
	  tray.Text = "SysCommander v1.0"

   End Sub

   ' ----------------------------------------------------------------------------------------------
   ' Menu handler

   Private Sub LaunchShellCommands(app As String)
	  Try
		 Process.Start("explorer.exe", "shell:" & app)
	  Catch ex As Exception
	  End Try
   End Sub

   Private Sub AppAbout(sender As Object, e As EventArgs)
	  Dim frm As New frmAbout()
	  frm.ShowDialog()
   End Sub

   Private Sub AppRebuild(sender As Object, e As EventArgs)
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%USERPROFILE%\Contacts"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%USERPROFILE%\3D Objects"))
   End Sub

   Private Sub AppExit(sender As Object, e As EventArgs)
	  tray.Visible = False
	  tray.Dispose()
	  ExitThread()
   End Sub

End Class
