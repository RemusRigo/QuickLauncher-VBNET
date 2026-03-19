'--------------------------------------------------------------------------------------------------
' QuickLauncher - A system tray application to quickly launch various shell folders and commands.
'    (C) 2026 Remus Rigo
'       v1.0.20260317
'--------------------------------------------------------------------------------------------------

Module SysTrayMain
   Sub Main()
      Application.EnableVisualStyles()
      Application.SetCompatibleTextRenderingDefault(False)
      Application.Run(New SysTray())
   End Sub

End Module
