Imports SMCRecover.ChunkModule

Module MainModule
    Private Declare Function SetConsoleCtrlHandler Lib "kernel32.dll" (ByVal HandlerRoutine As ConsoleCtrlHandler, ByVal Add As Boolean) As Boolean
    Private Delegate Function ConsoleCtrlHandler(evt As Integer) As Boolean
    Private dlg_Open As New Windows.Forms.OpenFileDialog
    Private dlg_Save As New Windows.Forms.SaveFileDialog
    Private Handler As ConsoleCtrlHandler

    Public MPQFIles As New Dictionary(Of String, MPQFile)
    Public CompLevel As MPQ_WaveLevel = MPQ_WaveLevel.Low
    Public Const SFMpq As String = "SFMpq.dll"
    Public Const Storm As String = "Storm.dll"
    Public Logger As IO.StreamWriter
    Public ChunkFile As ChunkData
    Public LastPath As String

    Public Structure MPQFile
        Dim FileData() As Byte
        Dim Compressions As Byte
        Dim IsWavFile As Boolean
        Dim IsEncrypted As Boolean

        Sub New(ByteData() As Byte, Optional Comp As Byte = 0, _
            Optional WavFile As Boolean = False, Optional Encrypted As Boolean = False)
            FileData = ByteData
            Compressions = Comp
            IsWavFile = WavFile
            IsEncrypted = Encrypted
        End Sub
    End Structure

    Sub Main()
        Dim OpenFile As String, SaveFile As String, LogFile As String
        Dim Args = My.Application.CommandLineArgs
        OpenFile = ""
        SaveFile = ""
        LogFile = ""

        '파라미터
        For i = 0 To Args.Count - 1
            If Args(i).ToLower() = "-output" Then
                i += 1
                LogFile = Args(i)
                Continue For
            End If
            If Args(i).ToLower() = "-open" Then
                i += 1
                OpenFile = Args(i)
                Continue For
            End If
            If Args(i).ToLower() = "-save" Then
                i += 1
                SaveFile = Args(i)
                Continue For
            End If
            If Args(i).ToLower() = "-comp" Then
                i += 1
                Select Case Args(i).ToLower()
                    Case "low", "l", 1
                        CompLevel = MPQ_WaveLevel.Low
                    Case "Medium", "m", 2
                        CompLevel = MPQ_WaveLevel.Medium
                    Case "high", "h", 0
                        CompLevel = MPQ_WaveLevel.High
                    Case Else
                        CompLevel = 4
                End Select
                Continue For
            End If
            If OpenFile.Length = 0 And IO.File.Exists(Args(i)) Then OpenFile = Args(i)
            If OpenFile.Length > 0 Then
                Try
                    SaveFile = IO.Path.GetFullPath(Args(i))
                Catch
                End Try
            End If
        Next

        '초기화
#If DllImport Then
        Call InitDll()
#End If
        Call InitConsole()
        Call InitLog(LogFile)
        WriteConsole("Open Path : " & IIf(OpenFile.Length = 0, "None", OpenFile))
        WriteConsole("Save Path : " & IIf(SaveFile.Length = 0, "None", SaveFile))
        If LogFile.Length > 0 Then WriteConsole("Log Path : " & LogFile)
        WriteConsole("Wav Compressions : " & CompLevel.ToString())

        '작업 시작
        If Not OpenMap(OpenFile) Then Return
        If Not Unprotect(ChunkFile) Then Return
        If Not SaveMap(SaveFile) Then Return
    End Sub

    Private Function ConsoleCtrlEvent(evt As Integer) As Boolean
        '콘솔 종료 이벤트
        If evt = 2 Then
            Logger.WriteLine()
            Logger.WriteLine("End Time : " & Now.ToString("yyyy\-MM\-dd ddd HH:MM:ss"))
            Logger.WriteLine("//--------------------------------------//")
            Logger.Close()
        End If
        Return True
    End Function

    Public Function OpenMap(FilePath As String) As Boolean
        Dim ByteData() As Byte, FileName As String, i As Integer
        If FilePath.Length = 0 Then
            dlg_Open.Filter = "Starcraft Map Files|*.scm; *.scx; *.cpg; *.chk|" & _
                "Starcraft Original Map|*.scm|Starcraft Expention Map|*.scx|" & _
                "Starcraft Campagin Map|*.cpg|Map Chunk Data|*.chk|All Files|*.*"
            dlg_Open.InitialDirectory = Environment.CurrentDirectory
            dlg_Open.FileName = ""
            If Not dlg_Open.ShowDialog() = Windows.Forms.DialogResult.OK Then Return False
            FileName = dlg_Open.FileName
        Else
            FileName = FilePath
        End If
        LastPath = FileName
        WriteConsole("Open Map File")
        WriteConsole("""" & FileName & """")
        MPQFIles.Clear()

        Select Case IO.Path.GetExtension(FileName).ToLower()
            Case ".scm", ".scx", ".cpg", ".mpq", ".exe"
                Dim hMPQ As Integer
                hMPQ = MPQRead_OpenMPQ(FileName)
                If hMPQ = -1 Then '올바르지 않은 MPQ 파일
                    WriteConsole("Invalid Map File!")
                    WriteConsole("FIle : " & FileName)
                    Return False
                End If

                'Chk파일 추출
                ByteData = MPQRead_ReadFile(hMPQ, ChunkPath)
                If ByteData Is Nothing Then         'Chk 파일이 없음
                    WriteConsole("Invalid Map File!")
                    WriteConsole("FIle : " & FileName)
                    MPQRead_CloseMPQ(hMPQ)
                    Return False
                End If

                'Chk 파일 분석
                ChunkFile = New ChunkData(ByteData)
                If ChunkFile.IndexOf("VER ") = -1 Or 0 <= ChunkFile.IndexOf("SMLP") Then
                    WriteConsole("Invalid Map File!")
                    WriteConsole("FIle : " & FileName)
                    MPQRead_CloseMPQ(hMPQ)
                    Return False
                End If
                WriteConsole("Chunk. " & ChunkFile.Count & "Sections " & GetFileSize(ByteData.Length))
                For Each Section In ChunkFile
                    WriteConsole(i & ". Section """ & Section.Name & """ Size. " & GetFileSize(Section.Data.Length))
                    i += 1
                Next

                'Wav 파일 추출
                For Each File In ChunkFile.GetWavStrings().Values
                    If MPQFIles.ContainsKey(File) Then Continue For
                    ByteData = MPQRead_ReadWavFile(hMPQ, File)
                    If ByteData Is Nothing Then Continue For
                    MPQFIles(File) = New MPQFile(ByteData, , True)
                    WriteConsole("Wav File. """ & File & """ Size. " & GetFileSize(ByteData.Length))
                Next
                MPQRead_CloseMPQ(hMPQ)

            Case ".chk"
                'Chk 파일 읽기
                Try
                    ByteData = IO.File.ReadAllBytes(FileName)
                Catch
                    WriteConsole("Cannot open file")
                    WriteConsole("FIle : " & FileName)
                    Return False
                End Try

                'Chk 파일 추출
                ChunkFile = New ChunkData(ByteData)
                If ChunkFile.IndexOf("VER ") = -1 Or 0 <= ChunkFile.IndexOf("SMLP") Then
                    WriteConsole("Invalid Map File!")
                    WriteConsole("FIle : " & FileName)
                    Return False
                End If
                WriteConsole("Chunk. " & ChunkFile.Count & "Sections " & GetFileSize(ByteData.Length))
                For Each Section In ChunkFile
                    WriteConsole(i & ". Section """ & Section.Name & """ Size. " & GetFileSize(Section.Data.Length))
                    i += 1
                Next

            Case Else
                WriteConsole("Unknown Map File!")
                WriteConsole("FIle : " & FileName)
                Return False
        End Select
        If ChunkFile.IndexOf("INFO") = -1 Then
            WriteConsole("Map not protected")
            Return False
        End If
        Return True
    End Function

    Public Function SaveMap(FilePath As String) As Boolean
        Dim ByteData() As Byte, FileName As String, i As Integer
        If FilePath.Length = 0 Then
            If ChunkFile.Version = &HCD Then
                dlg_Save.Filter = "Starcraft Expention Map|*.scx|Starcraft Campagin Map|*.cpg|Map Chunk Data|*.chk"
                dlg_Save.DefaultExt = ".scx"
            Else
                dlg_Save.Filter = "Starcraft Original Map|*.scm|Starcraft Expention Map|*.scx|Starcraft Campagin Map|*.cpg|Map Chunk Data|*.chk"
                dlg_Save.DefaultExt = ".scx"
            End If
            dlg_Save.InitialDirectory = IO.Path.GetDirectoryName(LastPath)
            dlg_Save.FileName = IO.Path.GetFileName(LastPath)
            If Not dlg_Save.ShowDialog() = Windows.Forms.DialogResult.OK Then Return False
            FileName = dlg_Save.FileName
        Else
            FileName = FilePath
        End If

        WriteConsole("Save Map File")
        WriteConsole("""" & FileName & """")
        If IO.File.Exists(FileName) Then
            Try
                IO.File.Delete(FileName)
            Catch
                WriteConsole("Invalid Path!")
                WriteConsole("FIle : " & FileName)
                Return False
            End Try
        End If
        Select Case IO.Path.GetExtension(FileName).ToLower()
            Case ".scm", ".scx", ".cpg", ".mpq"
                Dim Listfile As New List(Of String), hMPQ As Integer
                hMPQ = MPQWrite_OpenMPQ(FileName, MPQ_OpenType.Create)
                If hMPQ = -1 Then
                    WriteConsole("Invalid Path!")
                    WriteConsole("FIle : " & FileName)
                    Return False
                End If
                ByteData = ChunkFile.toArray()

                'MPQ에 파일 넣기
                Listfile.Add(ChunkPath)
                MPQWrite_WriteFile(hMPQ, ChunkPath, ByteData)
                If CompLevel = 4 Then
                    For Each File In MPQFIles
                        MPQWrite_WriteFile(hMPQ, File.Key, File.Value.FileData)
                        Listfile.Add(File.Key)
                    Next
                Else
                    For Each File In MPQFIles
                        MPQWrite_WriteWavFile(hMPQ, File.Key, File.Value.FileData, CompLevel)
                        Listfile.Add(File.Key)
                    Next
                End If
                MPQWrite_WriteFile(hMPQ, "(listfile)", System.Text.Encoding.Default.GetBytes(Join(Listfile.ToArray(), vbCrLf)))
                MPQWrite_CloseMPQ(hMPQ)

            Case ".chk"
                ByteData = ChunkFile.toArray()
                Try
                    IO.File.WriteAllBytes(FileName, ByteData)
                Catch
                    WriteConsole("Invalid Path!")
                    WriteConsole("FIle : " & FileName)
                    Return False
                End Try

            Case Else
                WriteConsole("Unknown Map File!")
                WriteConsole("FIle : " & FileName)
                Return False
        End Select
        WriteConsole("Chunk. " & ChunkFile.Count & "Sections " & GetFileSize(ByteData.Length))
        For Each Section In ChunkFile
            WriteConsole(i & ". Section """ & Section.Name & """ Size. " & GetFileSize(Section.Data.Length))
            i += 1
        Next
        WriteConsole("Save Map Succesfuly!")
        Return True
    End Function

#If DllImport Then
    Private Sub initdll()
        If Not IO.File.Exists(SFMpq) Then IO.File.WriteAllBytes(My.Application.Info.DirectoryPath & "\" & SFMpq, My.Resources.SFMpq)
        If Not IO.File.Exists(Storm) Then IO.File.WriteAllBytes(My.Application.Info.DirectoryPath & "\" & Storm, My.Resources.Storm)
    End Sub
#End If

    Private Sub InitConsole()
        Handler = New ConsoleCtrlHandler(AddressOf ConsoleCtrlEvent)
        SetConsoleCtrlHandler(Handler, True)
        Console.Title = "SMC Recover"
    End Sub

    Private Sub InitLog(Optional LogName As String = "")
        Dim FileName As String, NewName As String
        If LogName.Length = 0 Then
            FileName = "\Logfile-" & Process.GetCurrentProcess().Id & ".log"
            If IO.File.Exists(My.Application.Info.DirectoryPath & FileName) Then
                NewName = "\Recover " & IO.File.GetCreationTime(FileName).ToString("yyyy\-MM\-dd") & ".log"
                IO.File.Move(FileName, NewName)
            End If
            LogName = FileName
        End If
        Logger = New IO.StreamWriter(My.Application.Info.DirectoryPath & LogName)
        Logger.AutoFlush = True
        Logger.WriteLine("//-----------Recover Message Log-----------//")
        Logger.WriteLine("OS       : " & My.Computer.Info.OSFullName)
        Logger.WriteLine("Time     : " & Now.ToString("yyyy\-MM\-dd ddd HH:MM:ss"))
        Logger.WriteLine()
    End Sub

    Public Function GetFileSize(Size As Integer) As String
        If Size > &H100000 Then Return Size \ &H100000 & " MB"
        If Size > &H2800 Then Return Size \ &H400 & " KB"
        Return Size & "Bytes"
    End Function

    Public Sub WriteConsole(Msg As String)
        Console.WriteLine(Msg)
        Logger.WriteLine(Msg)
    End Sub
End Module
