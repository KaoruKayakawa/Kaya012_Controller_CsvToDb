Imports System.Diagnostics
Imports System.Threading
Imports System.Xml

Module Module1
    Private _s_exe As Semaphore
    Private _dic_proc As Dictionary(Of Integer, String)

    ' 多重起動を可とする
#If False Then
    Sub Main()
        Console.WriteLine("CSV インポート コントローラー")
        Console.WriteLine("")

        Console.WriteLine("■ アプリケーション多重起動の確認中　…")

        ' Global Mutex
        Dim mutexName As String = "Global\" + AppDomain.CurrentDomain.SetupInformation.ApplicationName.Split(".")(0)

        ' 'すべてのユーザーにフルコントロールを許可する MutexSecurity を作成
        Dim rule As New Security.AccessControl.MutexAccessRule(
            New Security.Principal.SecurityIdentifier(Security.Principal.WellKnownSidType.WorldSid, Nothing),
            Security.AccessControl.MutexRights.FullControl,
            Security.AccessControl.AccessControlType.Allow)
        Dim mutexSecurity As New Security.AccessControl.MutexSecurity()
        mutexSecurity.AddAccessRule(rule)

        Dim mutex As New Threading.Mutex(False, mutexName)
        mutex.SetAccessControl(mutexSecurity)

        Dim hasHandle As Boolean = False
        Try
            Try
                hasHandle = mutex.WaitOne(0, False)
            Catch ex As Threading.AbandonedMutexException       ' 別のアプリケーションがミューテックスを解放しないで終了
                hasHandle = True
            End Try

            If hasHandle = False Then
                Console.WriteLine("")
                Console.WriteLine("…　アプリケーションは実行中です。多重起動はできません。")
                Console.WriteLine("")

                Return
            End If

            MainSub1()
        Finally
            If hasHandle Then
                mutex.ReleaseMutex()
            End If
            mutex.Close()
        End Try
    End Sub

    Private Sub MainSub1()
        Console.WriteLine("■ 登録アプリケーションの順次実行を開始　…")

        _s_exe = New Semaphore(SettingConfig.ExeApps_MaxParallel, SettingConfig.ExeApps_MaxParallel)
        _dic_proc = New Dictionary(Of Integer, String)()

        For Each app As XmlNode In SettingConfig.ExeApps_App
            _s_exe.WaitOne()

            Console.WriteLine("-> " + SettingConfig.App_Name(app) + "　開始　（" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "）")

            Dim p As Process = New Process()
            p.StartInfo.FileName = SettingConfig.App_Path(app)
            p.StartInfo.UseShellExecute = True
            p.StartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(p.StartInfo.FileName)
            p.EnableRaisingEvents = True
            AddHandler p.Exited, New EventHandler(AddressOf proc_exited)

            p.Start()
            _dic_proc.Add(p.Id, SettingConfig.App_Name(app))
        Next

        Do While _dic_proc.Count > 0
            Thread.Sleep(1000)
        Loop

        Console.WriteLine("")
        Console.WriteLine("…　登録アプリケーションの処理が終了しました。")
        Console.WriteLine("")
    End Sub
#End If

    Sub Main()
        Console.WriteLine("CSV インポート コントローラー")
        Console.WriteLine("")

        Console.WriteLine("■ 登録アプリケーションの順次実行を開始　…")

        _s_exe = New Semaphore(SettingConfig.ExeApps_MaxParallel, SettingConfig.ExeApps_MaxParallel)
        _dic_proc = New Dictionary(Of Integer, String)()

        For Each app As XmlNode In SettingConfig.ExeApps_App
            _s_exe.WaitOne()

            Console.WriteLine("-> " + SettingConfig.App_Name(app) + "　開始　（" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "）")

            Dim p As Process = New Process()
            p.StartInfo.FileName = SettingConfig.App_Path(app)
            p.StartInfo.UseShellExecute = True
            p.StartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(p.StartInfo.FileName)
            p.EnableRaisingEvents = True
            AddHandler p.Exited, New EventHandler(AddressOf proc_exited)

            p.Start()
            _dic_proc.Add(p.Id, SettingConfig.App_Name(app))
        Next

        Do While _dic_proc.Count > 0
            Thread.Sleep(1000)
        Loop

        Console.WriteLine("")
        Console.WriteLine("…　登録アプリケーションの処理が終了しました。")
        Console.WriteLine("")
    End Sub

    Private Sub proc_exited(ByVal sender As Object, ByVal e As EventArgs)
        _s_exe.Release()

        Dim p As Process = DirectCast(sender, Process)
        Console.WriteLine("-> " + _dic_proc(p.Id) + "　終了　（" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "）")
        _dic_proc.Remove(p.Id)
    End Sub

End Module
