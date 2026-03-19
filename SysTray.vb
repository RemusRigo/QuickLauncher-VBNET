'--------------------------------------------------------------------------------------------------
' QuickLauncher - A system tray application to quickly launch various shell folders and commands.
'    (C) 2026 Remus Rigo
'       v1.0.20260319
'--------------------------------------------------------------------------------------------------

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

	  '----------------------------------------------------------------------------------------------
	  ' build left click menu

	  mnuLeft = New ContextMenuStrip()

	  mnuLeft.Items.Add("Rebuild missing folders", Nothing, AddressOf AppRebuild)
	  mnuLeft.Items.Add(New ToolStripSeparator())
	  mnuLeft.Items.Add("&About", Nothing, AddressOf AppAbout)
	  mnuLeft.Items.Add("E&xit", Nothing, AddressOf AppExit)

	  '----------------------------------------------------------------------------------------------
	  ' build right click menu

	  mnuRight = New ContextMenuStrip()

	  ' Shell Commands ----------------------------------------------------------------------------

	  Dim mnuShellCommands As New ToolStripMenuItem("Shell: commands")

	  Dim dicShellCommandsCategorized As New Dictionary(Of String, Dictionary(Of String, String)) From
	  {
		 {"Current User", New Dictionary(Of String, String) From
		 {
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
			{"Screenshots", "Shell:Screenshots"},
			{"Videos", "Shell:My Video"},
			{"Original Images", "Shell:Original Images"}, '
			{"Recycle Bin Folder", "Shell:RecycleBinFolder"},
			{"Ringtones", "Shell:Ringtones"},
			{"Saved Games", "Shell:SavedGames"},
			{"Searches", "Shell:Searches"},
			{"SendTo", "Shell:SendTo"},
			{"Libraries", "Shell:Libraries"},
			{"Documents Library", "Shell:DocumentsLibrary"},
			{"Music Library", "Shell:MusicLibrary"},
			{"Pictures Library", "Shell:PicturesLibrary"},
			{"Videos Library", "Shell:VideosLibrary"},
			{"Users Libraries Folder", "Shell:UsersLibrariesFolder"},
			{"User Pinned", "Shell:User Pinned"},
			{"User Profiles", "Shell:UserProfiles"},
			{"User ProgramFiles", "Shell:UserProgramFiles"},
			{"User ProgramFiles Common", "Shell:UserProgramFilesCommon"},
			{"User Tiles", "Shell:UserTiles"},
			{"Users Files Folder", "Shell:UsersFilesFolder"}
		 }
		 },
		 {"Current User AppData", New Dictionary(Of String, String) From
		 {
			{"AppData", "Shell:AppData"},
			{"Local AppData", "Shell:Local AppData"},
			{"Local AppData Low", "Shell:LocalAppDataLow"},
			{"Account Pictures", "Shell:AccountPictures"},
			{"Application Shortcuts", "Shell:Application Shortcuts"},
			{"Cache", "Shell:Cache"},
			{"CD Burning", "Shell:CD Burning"},
			{"Cookies", "Shell:Cookies"},
			{"Cookies\Low", "Shell:Cookies\Low"},
			{"Credential Manager", "Shell:CredentialManager"},
			{"Cryptokeys", "Shell:Cryptokeys"},
			{"Dp API Keys", "Shell:DpAPIKeys"},
			{"History", "Shell:History"},
			{"Gadgets (Win 7)", "Shell:Gadgets"}, 'Windows 7 only
			{"GameTasks", "Shell:GameTasks"},
			{"Implicit App Shortcuts", "Shell:ImplicitAppShortcuts"},
			{"Quick Launch", "Shell:Quick Launch"},
			{"PrintHood", "Shell:PrintHood"},
			{"Recent", "Shell:Recent"},
			{"Start Menu", "Shell:Start Menu"},
			{"Startup", "Shell:Startup"},
			{"System Certificates", "Shell:SystemCertificates"},
			{"Templates", "Shell:Templates"}
		 }
		 },
		{"Public", New Dictionary(Of String, String) From
		{
			{"Common Desktop", "Shell:Common Desktop"},
			{"Common Documents", "Shell:Common Documents"},
			{"Common Downloads", "Shell:CommonDownloads"},
			{"Common Music", "Shell:CommonMusic"},
			{"Common Pictures", "Shell:CommonPictures"},
			{"Common Video", "Shell:CommonVideo"},
			{"Common Start Menu", "Shell:Common Start Menu"},
			{"Common Startup", "Shell:Common Startup"},
			{"Public", "Shell:Public"},
			{"Public Account Pictures", "Shell:PublicAccountPictures"},
			{"Public Libraries", "Shell:PublicLibraries"},
			{"Public User Tiles", "Shell:PublicUserTiles"}
		 }
		 },
		{"Control Panel", New Dictionary(Of String, String) From
		{
		   {"Control Panel Folder", "Shell:ControlPanelFolder"},
		   {"Add New Programs Folder", "Shell:AddNewProgramsFolder"},
		   {"Administrative Tools", "Shell:Administrative Tools"},
		   {"Apps Folder", "Shell:AppsFolder"},
		   {"Change Remove Programs Folder", "Shell:ChangeRemoveProgramsFolder"},
		   {"Connections Folder", "Shell:ConnectionsFolder"},
		   {"NetHood", "Shell:NetHood"},
		   {"Printers Folder", "Shell:PrintersFolder"},
		   {"SyncCenter Folder", "Shell:SyncCenterFolder"},
		   {"SyncCenter Results Folder", "Shell:SyncResultsFolder"},
		   {"SyncCenter Setup Folder", "Shell:SyncSetupFolder"},
		   {"SyncCenter Conflict Folder", "Shell:ConflictFolder"},
		   {"SyncCenter CSC (Client Side Caching) Folder", "Shell:CSCFolder"},
		   {"Windows", "Shell:Windows"}
		}
		},
		{"Settings", New Dictionary(Of String, String) From
		{
		   {"App Updates Folder", "Shell:AppUpdatesFolder"}
		}
		},
		{"Windows", New Dictionary(Of String, String) From
		{
		   {"Fonts", "Shell:Fonts"},
		   {"My Computer Folder", "Shell:MyComputerFolder"},
		   {"Network Places Folder", "Shell:NetworkPlacesFolder"},
		   {"Resource Dir", "Shell:ResourceDir"},
		   {"System (System32)", "Shell:System"},
		   {"System X86 (SysWOW64)", "Shell:SystemX86"},
		   {"Windows", "Shell:Windows"}
		}
		},
		{"Program Files", New Dictionary(Of String, String) From
		{
		   {"ProgramFiles", "Shell:ProgramFiles"},
		   {"ProgramFiles Common", "Shell:ProgramFilesCommon"},
		   {"ProgramFiles CommonX64", "Shell:ProgramFilesCommonX64"},
		   {"ProgramFiles CommonX86", "Shell:ProgramFilesCommonX86"},
		   {"ProgramFiles X64", "Shell:ProgramFilesX64"},
		   {"ProgramFiles X86", "Shell:ProgramFilesX86"},
		   {"Programs", "Shell:Programs"},
		   {"Default Gadgets (Win Vista/7/8)", "Shell:Default Gadgets"} 'Windows Vista, Windows 7, and Windows 8		
		}
		},
		{"ProgramData", New Dictionary(Of String, String) From
		{
		   {"Common AppData", "Shell:Common AppData"},
		   {"Common Administrative Tools", "Shell:Common Administrative Tools"},
		   {"Common Programs", "Shell:Common Programs"},
		   {"Common Ringtones", "Shell:CommonRingtones"},
		   {"Common Templates", "Shell:Common Templates"},
		   {"Device Metadata Store", "Shell:Device Metadata Store"},
		   {"OEM Links", "Shell:OEM Links"},
		   {"Public Game Tasks", "Shell:PublicGameTasks"}
		}
		},
		{"OneDrive", New Dictionary(Of String, String) From
		{
		   {"OneDrive", "Shell:OneDrive"},
		   {"OneDrive CameraRoll", "Shell:OneDriveCameraRoll"},
		   {"OneDrive Documents", "Shell:OneDriveDocuments"},
		   {"OneDrive Music", "Shell:neDriveMusic"},
		   {"OneDrive Pictures", "Shell:OneDrivePictures"},
		   {"SkyDrive CameraRoll ", "Shell:SkyDriveCameraRoll"},
		   {"SkyDrive Documents ", "Shell:SkyDriveDocuments"},
		   {"SkyDrive Music ", "Shell:SkyDriveMusic"},
		   {"SkyDrive Pictures ", "Shell:SkyDrivePictures"}
		}
		}
	  }

	  Dim dicShellCommandsUncategorized As New Dictionary(Of String, String) From
	  {
		 {"HomeGroup Current User Folder", "Shell:HomeGroupCurrentUserFolder"}, '
		 {"HomeGroup Folder", "Shell:HomeGroupFolder"}, '
		 {"Internet Folder", "Shell:InternetFolder"}, ' deprecated		 
		 {"Localized Resources Dir", "Shell:LocalizedResourcesDir"}, '
		 {"MAPI Folder ", "Shell:MAPIFolder"}, '		 		 		
		 {"Recorded TV Library", "Shell:RecordedTVLibrary"},
		 {"Retail Demo", "Shell:Retail Demo"}, '
		 {"Roamed Tile Images", "Shell:Roamed Tile Images"},
		 {"Roaming Tiles", "Shell:Roaming Tiles"},
		 {"Sample Music", "Shell:SampleMusic"},
		 {"Sample Pictures", "Shell:SamplePictures"},
		 {"Sample Videos", "Shell:SampleVideos"},
		 {"Search History Folder", "Shell:SearchHistoryFolder"},
		 {"Search Home Folder", "Shell:SearchHomeFolder"},
		 {"Search Templates Folder", "Shell:SearchTemplatesFolder"},
		 {"Start Menu All Programs", "Shell:StartMenuAllPrograms"},
		 {"This PC Desktop Folder", "Shell:ThisPCDesktopFolder"}
	  }

	  For Each category In dicShellCommandsCategorized.Keys
		 Dim categoryItem As New ToolStripMenuItem(category)
		 For Each entry In dicShellCommandsCategorized(category)
			Dim displayName As String = entry.Key
			Dim shellName As String = entry.Value
			Dim item As New ToolStripMenuItem(displayName)
			Dim localShell = shellName
			AddHandler item.Click, Sub() LaunchShellCommands(localShell)
			categoryItem.DropDownItems.Add(item)
		 Next
		 mnuShellCommands.DropDownItems.Add(categoryItem)
	  Next

	  For Each entry In dicShellCommandsUncategorized
		 Dim item As New ToolStripMenuItem(entry.Key)
		 Dim localShell = entry.Value
		 AddHandler item.Click, Sub() LaunchShellCommands(localShell)
		 mnuShellCommands.DropDownItems.Add(item)
	  Next

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

	Private Sub AppAbout(sender As Object, e As EventArgs)
		Using frm As New frmAbout()
			frm.ShowInTaskbar = False
			frm.ShowDialog()
			' Dispose() is called automatically when exiting the Using block, even if an exception occurs
		End Using
	End Sub

	Private Sub AppExit(sender As Object, e As EventArgs)
		tray.Visible = False
		tray.Dispose()
		ExitThread()
	End Sub

End Class
