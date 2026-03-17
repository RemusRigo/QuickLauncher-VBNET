Imports System.Diagnostics.Contracts
Imports System.IO
Imports System.Net
Imports System.Net.Mime.MediaTypeNames
Imports System.Reflection
Imports System.Reflection.Metadata
Imports System.Resources
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.JavaScript.JSType
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.VisualBasic.Logging

Public Class SysTray
   Inherits ApplicationContext

   Private tray As NotifyIcon
   Private mnuLeft As ContextMenuStrip
   Private mnuRight As ContextMenuStrip

   <DllImport("user32.dll")>
   Private Shared Function GetDesktopWindow() As IntPtr
   End Function

   <DllImport("user32.dll")>
   Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
   End Function


   Public Sub New()
	  ' Build menu
	  mnuLeft = New ContextMenuStrip()
	  mnuRight = New ContextMenuStrip()

	  ' Shell Commands ----------------------------------------------------------------------------

	  Dim mnuShellCommands As New ToolStripMenuItem("Shell: commands")

	  Dim dicShellFolderscategorized As New Dictionary(Of String, Dictionary(Of String, String)) From {
		 {"Current User", New Dictionary(Of String, String) From {
			{"Profile", "Shell:Profile"},
			{"3D Objects", "Shell:3D Objects"},
			{"Contacts", "Shell:Contacts"},
			{"Desktop", "Shell:Desktop"},
			{"Downloads", "Shell:Downloads"},
			{"Favorites", "Shell:Favorites"},
			{"Games", "Shell:Games"},
			{"Links", "Shell:Links"},
			{"Music", "Shell:My Music"},
			{"Playlists", "Shell:Playlists"},
			{"Personal", "Shell:Personal"},
			{"Pictures", "Shell:My Pictures"},
			{"Camera Roll", "Shell:Camera Roll"},
			{"Photo Albums", "Shell:PhotoAlbums"},
			{"Screenshots", "Screenshots"},
			{"Videos", "My Video"},
			{"Original Images", "Original Images"}, '
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
			{"Gadgets (Win 7)", "Gadgets"}, 'Windows 7 only
			{"GameTasks", "GameTasks"},
			{"Implicit App Shortcuts", "ImplicitAppShortcuts"},
			{"Quick Launch", "Quick Launch"},
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
		   {"Resource Dir", "ResourceDir"},
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
		   {"Programs", "Programs"},
		   {"Default Gadgets (Win Vista/7/8)", "Default Gadgets"} 'Windows Vista, Windows 7, and Windows 8		
		}},
		{"ProgramData", New Dictionary(Of String, String) From {
		   {"Common AppData", "Common AppData"},
		   {"Common Administrative Tools", "Common Administrative Tools"},
		   {"Common Programs", "Common Programs"},
		   {"Common Ringtones", "CommonRingtones"},
		   {"Common Templates", "Common Templates"},
		   {"Device Metadata Store", "Device Metadata Store"},
		   {"OEM Links", "OEM Links"},
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
		 {"HomeGroup Current User Folder", "HomeGroupCurrentUserFolder"}, '
		 {"HomeGroup Folder", "HomeGroupFolder"}, '
		 {"Internet Folder", "InternetFolder"}, ' deprecated		 
		 {"Localized Resources Dir", "LocalizedResourcesDir"}, '
		 {"MAPI Folder ", "MAPIFolder"}, '		 		 		
		 {"Recorded TV Library", "RecordedTVLibrary"},
		 {"Retail Demo", "Retail Demo"}, '
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
		 mnuShellCommands.DropDownItems.Add(categoryItem)
	  Next

	  For Each entry In dicShellFoldersUncategorized
		 Dim item As New ToolStripMenuItem(entry.Key)
		 Dim localShell = entry.Value
		 AddHandler item.Click, Sub() LaunchShellCommands(localShell)
		 mnuShellCommands.DropDownItems.Add(item)
	  Next

	  ' Left-click menu
	  mnuLeft.Items.Add("Rebuild missing folders", Nothing, AddressOf AppRebuild)
	  mnuLeft.Items.Add(New ToolStripSeparator())
	  mnuLeft.Items.Add("&About", Nothing, AddressOf AppAbout)
	  mnuLeft.Items.Add("E&xit", Nothing, AddressOf AppExit)

	  ' Right-click menu
	  mnuRight.Items.Add(mnuShellCommands)
	  mnuRight.Items.Add(New ToolStripSeparator())
	  mnuRight.Items.Add("&About", Nothing, AddressOf AppAbout)
	  mnuRight.Items.Add("E&xit", Nothing, AddressOf AppExit)

	  ' Create tray icon
	  tray = New NotifyIcon()
	  tray.Icon = My.Resources.Resources.terminal
	  tray.Visible = True
	  tray.ContextMenuStrip = mnuRight
	  tray.Text = "QuickLauncher v1.0"

	  AddHandler tray.MouseClick, AddressOf Tray_MouseClick
	  AddHandler tray.MouseDown, AddressOf Tray_MouseClick
	  AddHandler tray.MouseUp, AddressOf Tray_MouseClick
   End Sub


   Private Sub Tray_MouseClick(sender As Object, e As MouseEventArgs)
	  If e.Button = MouseButtons.Left Then
		 Dim mi As MethodInfo = GetType(NotifyIcon).GetMethod("ShowContextMenu", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
		 tray.ContextMenuStrip = mnuLeft
		 mi?.Invoke(tray, Nothing)
	  ElseIf e.Button = MouseButtons.Right Then
		 Dim mi As MethodInfo = GetType(NotifyIcon).GetMethod("ShowContextMenu", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
		 tray.ContextMenuStrip = mnuRight
		 mi?.Invoke(tray, Nothing)
	  End If
	  tray.ContextMenuStrip = Nothing
   End Sub

   ' ----------------------------------------------------------------------------------------------
   ' Menu handler

   Private Sub LaunchShellCommands(app As String)
	  Try
		 Process.Start("explorer.exe", app)
	  Catch ex As Exception
	  End Try
   End Sub

   Private Sub AppAbout(sender As Object, e As EventArgs)
	  Dim frm As New frmAbout()
	  frm.ShowInTaskbar = False
	  frm.ShowDialog()
	  frm.Dispose()
   End Sub

   Private Sub AppRebuild(sender As Object, e As EventArgs)
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%UserProfile%\3D Objects"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%UserProfile%\Contacts"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%UserProfile%\Music\Playlists"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%UserProfile%\Pictures\Camera Roll"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%UserProfile%\Pictures\Slide Shows"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%LocalAppData%\Microsoft\Windows\Burn\Burn"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%LocalAppData%\Microsoft\Windows Photo Gallery\Original Images"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%ProgramFiles%\Windows Sidebar\Gadgets"))
	  System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables("%ProgramData%\OEM Links"))
   End Sub

   Private Sub AppExit(sender As Object, e As EventArgs)
	  tray.Visible = False
	  tray.Dispose()
	  ExitThread()
   End Sub

End Class
