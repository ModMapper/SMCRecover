Imports System.Runtime.InteropServices
Imports System.IO

Namespace ChunkModule
    Partial Public Class ChunkData
        Public Property Version As UShort
            Get
                Dim Section As CSection
                Section = Sections(LastIndexOf("VER "))
                If Section Is Nothing Then Return &HFF
                Section.Data.Position = 0
                Return Section.Reader.ReadUInt16()
            End Get
            Set(Value As UShort)
                Dim Section As CSection
                Section = Sections(LastIndexOf("VER "))
                If Section Is Nothing Then Section = Sections(Add("VER "))
                Section.Data.Position = 0
                Section.Writer.Write(Value)
            End Set
        End Property

        Public Property Tileset As UShort
            Get
                Dim Section As CSection
                Section = Sections(LastIndexOf("ERA "))
                If Section Is Nothing Then Return 0
                Section.Data.Position = 0
                Return Section.Reader.ReadUInt16() And 7
            End Get
            Set(Value As UShort)
                Dim Section As CSection
                Section = Sections(LastIndexOf("ERA "))
                If Section Is Nothing Then Section = Sections(Add("ERA "))
                Section.Data.Position = 0
                Section.Writer.Write(Value)
            End Set
        End Property

        Public Property MapSize As MapSize
            Get
                Dim Section As CSection
                Section = Sections(LastIndexOf("DIM "))
                If Section Is Nothing Then Return Nothing
                Section.Data.Position = 0
                Return New MapSize(Section.Reader.ReadInt16(), Section.Reader.ReadInt16())
            End Get
            Set(Value As MapSize)
                Dim Section As CSection
                Section = Sections(LastIndexOf("DIM "))
                If Section Is Nothing Then Section = Sections(Add("DIM "))
                Section.Data.Position = 0
                Section.Writer.Write(CShort(Value.Width))
                Section.Writer.Write(CShort(Value.Height))
            End Set
        End Property

        Public Function GetString(StrIndex As Integer) As String
            Dim Section As CSection, Buffer As IntPtr
            Dim StrData As String
            Section = GetOverwrite("STR ")
            Buffer = Marshal.AllocCoTaskMem(Section.Data.Length)
            Marshal.Copy(Section.Data.ToArray(), 0, Buffer, Section.Data.Length)
            StrData = GetString(Buffer, Section.Data.Length, StrIndex)
            Marshal.FreeCoTaskMem(Buffer)
            Return StrData
        End Function

        Private Function GetString(Buffer As IntPtr, Size As Integer, StrIndex As Integer) As String
            Dim Index As UShort
            If Size <= StrIndex * 2 Then Return ""
            Index = Marshal.ReadInt16(Buffer, StrIndex * 2)
            If Size <= Index Then Return ""
            Return Marshal.PtrToStringAnsi(Buffer + Index)
        End Function


        Public Sub SetStrings(StrList As Dictionary(Of Integer, String), Optional StrSize As UShort = 1024)
            Dim NewIndex() As Integer, OldIndex() As Integer
            Dim Size As UShort, i As Integer
            Dim Section As New CSection
            Dim DataIndex() As UShort
            SectionList.Insert(IndexOf("STR "), Section)
            Remove("STR ")
            Section.Name = "STR "

            'String Write
            If StrList.ContainsKey(0) Then StrList.Remove(0)
            Size = Math.Max(StrList.Count, StrSize)
            NewIndex = Array.CreateInstance(GetType(Integer), StrList.Count)
            OldIndex = Array.CreateInstance(GetType(Integer), StrList.Count)
            DataIndex = Array.CreateInstance(GetType(UShort), Size)
            Section.Data.SetLength(Size * 2 + 2)
            Section.Data.Seek(0, SeekOrigin.End)
            Section.Data.WriteByte(0)
            For Each StrData In StrList
                DataIndex(i) = CUShort(Section.Data.Position)
                Section.Writer.Write(System.Text.Encoding.Default.GetBytes(StrData.Value))
                Section.Data.WriteByte(0)
                OldIndex(i) = StrData.Key
                NewIndex(i) = i + 1
                i += 1
            Next
            Do While i < Size
                DataIndex(i) = Size * 2
                i += 1
            Loop

            'Offset Write
            Section.Data.Position = 0
            Section.Writer.Write(Size)
            For i = 0 To Size - 1
                Section.Writer.Write(DataIndex(i))
            Next

            SetStrings(NewIndex, OldIndex)
        End Sub

        Friend Sub SetStrings(NewIndex() As Integer, OldIndex() As Integer)
            Dim Stream As New MemoryStream(), Reader As BinaryReader
            Dim Trigger() As Trigger, Location() As LocationData
            Dim GetIndex As Func(Of Integer, Integer)
            Dim i As Integer, n As Integer
            Dim Section As New CSection
            Dim Looped As Integer
            GetIndex = New Func(Of Integer, Integer)(
                Function(StrIndex As Integer)
                    Dim Index As Integer
                    Index = Array.IndexOf(Of Integer)(OldIndex, StrIndex)
                    If Index = -1 Then Return 0
                    Return NewIndex(Index)
                End Function)
            Reader = New BinaryReader(Stream)

            'Map Discription
            Section = Sections(IndexOf("SPRP"))
            Section.Data.Position = 0
            Section.Data.WriteTo(Stream)
            Stream.Position = 0
            For i = 0 To 1
                Section.Writer.Write(CUShort(GetIndex(Reader.ReadUInt16())))
            Next

            'Force
            Section = Sections(IndexOf("FORC"))
            Stream.SetLength(0)
            Section.Data.Position = 8
            Section.Data.WriteTo(Stream)
            Stream.Position = 8
            For i = 0 To 3
                Section.Writer.Write(CUShort(GetIndex(Reader.ReadUInt16())))
            Next

            'Locations
            Location = Locations()
            For i = 0 To Location.Length - 1
                Location(i).NameString = GetIndex(Location(i).NameString)
            Next
            Locations() = Location

            'Switchs, Wav Files
            Do
                Section = Sections(IndexOf(IIf(Looped = 0, "SWNM", "WAV ")))
                If Section IsNot Nothing Then
                    Stream.SetLength(0)
                    Section.Data.Position = 0
                    Section.Data.WriteTo(Stream)
                    Stream.Position = 0
                    For i = 0 To Section.Data.Length / 4 - 1
                        Section.Writer.Write(GetIndex(Reader.ReadUInt16()))
                    Next
                End If
                Looped += 1
            Loop Until Looped = 2
            Looped = 0

            'Units
            Do
                Section = Sections(IndexOf(IIf(Looped = 0, "UNIS", "UNIx")))
                If Section IsNot Nothing Then
                    Stream.SetLength(0)
                    Section.Data.WriteTo(Stream)
                    Section.Data.Position = &HC78
                    Stream.Position = &HC78
                    For i = 0 To 227
                        Section.Writer.Write(CUShort(GetIndex(Reader.ReadUInt16())))
                    Next
                End If
                Looped += 1
            Loop Until Looped = 2

            'Trigger
            Trigger = Triggers().ToArray()
            For i = 0 To Trigger.Length - 1
                For n = 0 To 63
                    Select Case Trigger(i).Actions(n).Action
                        Case TActionData.NoAction
                            Exit For
                        Case TActionData.PlayWav
                            Trigger(i).Actions(n).WavString = GetIndex(Trigger(i).Actions(n).WavString)
                        Case TActionData.Transmission
                            Trigger(i).Actions(n).WavString = GetIndex(Trigger(i).Actions(n).WavString)
                            Trigger(i).Actions(n).MsgString = GetIndex(Trigger(i).Actions(n).MsgString)
                        Case 9, 12, 17, 18, 19, 20, 21, 33, 34, 35, 36, 37, 41, 47
                            Trigger(i).Actions(n).MsgString = GetIndex(Trigger(i).Actions(n).MsgString)
                    End Select
                Next
            Next
            Triggers() = New List(Of Trigger)(Trigger)

            'Mission Brifing
            Trigger = MissionBrifing().ToArray()
            For i = 0 To Trigger.Length - 1
                For n = 0 To 63
                    Select Case Trigger(i).Actions(n).Action
                        Case MActionData.NoAction
                            Exit For
                        Case MActionData.PlayWav
                            Trigger(i).Actions(n).WavString = GetIndex(Trigger(i).Actions(n).WavString)
                        Case MActionData.Transmission
                            Trigger(i).Actions(n).WavString = GetIndex(Trigger(i).Actions(n).WavString)
                            Trigger(i).Actions(n).MsgString = GetIndex(Trigger(i).Actions(n).MsgString)
                        Case MActionData.TextMessage, MActionData.MissionObjectives
                            Trigger(i).Actions(n).MsgString = GetIndex(Trigger(i).Actions(n).MsgString)
                    End Select
                Next
            Next
            MissionBrifing() = New List(Of Trigger)(Trigger)
        End Sub

        Public Function GetStrings() As Dictionary(Of Integer, String)
            Dim StrList As New Dictionary(Of Integer, String)
            Dim StrIndex As Integer, i As Integer
            Dim Buffer As IntPtr, Size As Integer
            Dim Section As CSection
            Dim Looped As Integer
            'String Buffer
            Section = GetOverwrite("STR ")
            Size = Section.Data.Length
            Buffer = Marshal.AllocCoTaskMem(Size)
            Marshal.Copy(Section.Data.ToArray(), 0, Buffer, Size)
            StrList(0) = ""

            'Map Discription
            Section = Sections(IndexOf("SPRP"))
            Section.Data.Position = 0
            For i = 0 To 1
                StrIndex = Section.Reader.ReadUInt16()
                If Not StrList.ContainsKey(StrIndex) Then _
                    StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
            Next

            'Force
            Section = Sections(IndexOf("FORC"))
            Section.Data.Position = 8
            For i = 0 To 3
                StrIndex = Section.Reader.ReadUInt16()
                If Not StrList.ContainsKey(StrIndex) Then _
                    StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
            Next

            'Locations
            For Each Location In Locations()
                StrIndex = Location.NameString
                If Not StrList.ContainsKey(StrIndex) Then _
                    StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
            Next

            'Switchs, Wav Files
            Do
                Section = Sections(IndexOf(IIf(Looped = 0, "SWNM", "WAV ")))
                If Section IsNot Nothing Then
                    Section.Data.Position = 0
                    For i = 0 To Section.Data.Length / 4 - 1
                        StrIndex = Section.Reader.ReadInt32()
                        If Not StrList.ContainsKey(StrIndex) Then _
                            StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                    Next
                End If
                Looped += 1
            Loop Until Looped = 2
            Looped = 0

            'Units
            Do
                Section = Sections(IndexOf(IIf(Looped = 0, "UNIS", "UNIx")))
                If Section IsNot Nothing Then
                    Section.Data.Position = &HC78
                    For i = 0 To 227
                        StrIndex = Section.Reader.ReadUInt16()
                        If Not StrList.ContainsKey(StrIndex) Then _
                            StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                    Next
                End If
                Looped += 1
            Loop Until Looped = 2

            'Trigger
            Dim a As Integer
            For Each Trig In Triggers()
                For i = 0 To 63
                    Select Case Trig.Actions(i).Action
                        Case TActionData.NoAction
                            Exit For
                        Case TActionData.PlayWav
                            StrIndex = Trig.Actions(i).WavString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                        Case TActionData.Transmission
                            StrIndex = Trig.Actions(i).WavString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                            StrIndex = Trig.Actions(i).MsgString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                        Case 9, 12, 17, 18, 19, 20, 21, 33, 34, 35, 36, 37, 41, 47
                            StrIndex = Trig.Actions(i).MsgString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                    End Select
                Next
                a += 1
            Next

            'Mission Brifing
            For Each Trig In MissionBrifing()
                For i = 0 To 63
                    Select Case Trig.Actions(i).Action
                        Case MActionData.NoAction
                            Exit For
                        Case MActionData.PlayWav
                            StrIndex = Trig.Actions(i).WavString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                        Case MActionData.Transmission
                            StrIndex = Trig.Actions(i).WavString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                            StrIndex = Trig.Actions(i).MsgString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                        Case MActionData.TextMessage, MActionData.MissionObjectives
                            StrIndex = Trig.Actions(i).MsgString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                    End Select
                Next
            Next
            Marshal.FreeCoTaskMem(Buffer)
            Return StrList
        End Function

        Public Function GetWavStrings() As Dictionary(Of Integer, String)
            Dim StrList As New Dictionary(Of Integer, String)
            Dim Buffer As IntPtr, Size As Integer
            Dim StrIndex As Integer, i As Integer
            Dim Section As CSection
            'String Buffer
            Section = GetOverwrite("STR ")
            Size = Section.Data.Length
            Buffer = Marshal.AllocCoTaskMem(Size)
            Marshal.Copy(Section.Data.ToArray(), 0, Buffer, Size)
            StrList(0) = ""

            'Wav Files
            Section = Sections(IndexOf("WAV "))
            If Section IsNot Nothing Then
                Section.Data.Position = 0
                For i = 0 To Section.Data.Length / 4 - 1
                    StrIndex = Section.Reader.ReadInt32()
                    If Not StrList.ContainsKey(StrIndex) Then _
                        StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                Next
            End If

            'Trigger
            For Each Trig In Triggers()
                For i = 0 To 63
                    Select Case Trig.Actions(i).Action
                        Case TActionData.NoAction
                            Exit For
                        Case TActionData.PlayWav, TActionData.Transmission
                            StrIndex = Trig.Actions(i).WavString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                    End Select
                Next
            Next

            'Mission Brifing
            For Each Trig In MissionBrifing()
                For i = 0 To 63
                    Select Case Trig.Actions(i).Action
                        Case MActionData.NoAction
                            Exit For
                        Case MActionData.PlayWav, MActionData.Transmission
                            StrIndex = Trig.Actions(i).WavString
                            If Not StrList.ContainsKey(StrIndex) Then _
                                StrList(StrIndex) = GetString(Buffer, Size, StrIndex)
                    End Select
                Next
            Next
            Marshal.FreeCoTaskMem(Buffer)
            Return StrList
        End Function

        Public Property Units() As UnitData()
            Get
                Dim Section As CSection
                Section = Sections(IndexOf("UNIT"))
                If Section Is Nothing Then Return Nothing
                Return StructureArrayCreate(Of UnitData)(Section)
            End Get
            Set(Unit As UnitData())
                Dim Section As CSection
                If Unit Is Nothing Then Exit Property
                Section = Sections(IndexOf("UNIT"))
                If Section Is Nothing Then
                    Section = Sections(Add("UNIT"))
                End If
                StructureArrayWrite(Of UnitData)(Section, Unit)
            End Set
        End Property

        Public Property Doodads() As DoodadData()
            Get
                Dim Section As CSection
                Section = Sections(IndexOf("DD2 "))
                If Section Is Nothing Then Return Nothing
                Return StructureArrayCreate(Of DoodadData)(Section)
            End Get
            Set(Doodad As DoodadData())
                Dim Section As CSection
                If Doodad Is Nothing Then Exit Property
                Section = Sections(IndexOf("DD2 "))
                If Section Is Nothing Then
                    Section = Sections(Add("DD2 "))
                End If
                StructureArrayWrite(Of DoodadData)(Section, Doodad)
            End Set
        End Property

        Public Property Sprites() As SpriteData()
            Get
                Dim Section As CSection
                Section = Sections(IndexOf("THG2"))
                If Section Is Nothing Then Return Nothing
                Return StructureArrayCreate(Of SpriteData)(Section)
            End Get
            Set(Sprite As SpriteData())
                Dim Section As CSection
                If Sprite Is Nothing Then Exit Property
                Section = Sections(IndexOf("THG2"))
                If Section Is Nothing Then
                    Section = Sections(Add("THG2"))
                End If
                StructureArrayWrite(Of SpriteData)(Section, Sprite)
            End Set
        End Property

        Public Property Locations() As LocationData()
            Get
                Dim Section As CSection
                Section = Sections(IndexOf("MRGN"))
                If Section Is Nothing Then Return Nothing
                Return StructureArrayCreate(Of LocationData)(Section)
            End Get
            Set(Location As LocationData())
                Dim Section As CSection
                Section = Sections(IndexOf("MRGN"))
                If Section Is Nothing Then Section = Sections(Add("MRGN"))
                StructureArrayWrite(Of LocationData)(Section, Location)
            End Set
        End Property

        Public Property MissionBrifing() As List(Of Trigger)
            Get
                Return Triggers(True)
            End Get
            Set(TrigList As List(Of Trigger))
                Triggers(True) = TrigList
            End Set
        End Property

        Public Property Triggers(Optional MBRF As Boolean = False) As List(Of Trigger)
            Get
                Dim TrigList As New List(Of Trigger)
                Dim ByteData() As Byte, i As Integer
                For Each Section In Sections(CStr(IIf(MBRF, "MBRF", "TRIG")))
                    Section.Data.Position = 0
                    For i = 1 To Section.Data.Length Step 2400
                        ByteData = Section.Reader.ReadBytes(2400)
                        TrigList.Add(StructureCreate(Of Trigger)(ByteData))
                    Next
                Next
                Return TrigList
            End Get
            Set(TrigList As List(Of Trigger))
                Dim Name As String, Index As Integer
                Dim Section As New CSection
                Name = IIf(MBRF, "MBRF", "TRIG")
                Section.Name = Name
                Index = IndexOf(Name)
                Remove(Name)
                If Index = -1 Then
                    If TrigList.Count = 0 Then Exit Property
                    Index = Count
                End If
                SectionList.Insert(Index, Section)
                StructureArrayWrite(Of Trigger)(Section, TrigList.ToArray())
            End Set
        End Property
    End Class

    Public Module Starcraft
        Public ReadOnly SectionSort() As String = {"TYPE", "VER ", "IVER", "IVE2", "VCOD", "IOWN", "OWNR", "ERA ", "DIM ", "SIDE", _
        "MTXM", "PUNI", "UPGR", "PTEC", "UNIT", "ISOM", "TILE", "DOOD", "DD2 ", "THGY", "THG2", "MASK", "STR ", "UPRP", "UPUS", _
        "MRGN", "TRIG", "MBRF", "SPRP", "FORC", "WAV ", "UNIS", "UPGS", "TECS", "SWNM", "COLR", "PUPx", "PTEx", "UNIx", "UPGx", "TECx"}

        Public Structure MapSize
            Dim Width As Integer
            Dim Height As Integer

            Sub New(W As UShort, H As UShort)
                Width = W
                Height = H
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential)> Public Structure DoodadData   '8 Bytes
            Dim ID As UShort
            Dim X As UShort
            Dim Y As UShort
            Dim Player As Byte
            Dim Flags As Byte
        End Structure

        <StructLayout(LayoutKind.Sequential)> Public Structure SpriteData   '10 Bytes
            Dim ID As UShort
            Dim X As UShort
            Dim Y As UShort
            Dim Player As Byte
            Dim Unused As Byte
            Dim Flags As UShort
        End Structure

        <StructLayout(LayoutKind.Sequential)> Public Structure UnitData     '36 Bytes
            Dim ClassInstance As UInteger
            Dim X As UShort
            Dim Y As UShort
            Dim ID As UShort
            Dim LinkType As UShort
            Dim StateValid As UShort
            Dim DataValid As UShort
            Dim Player As Byte
            Dim HitPoint As Byte
            Dim ShieldPoint As Byte
            Dim EnergyPoint As Byte
            Dim Resource As UInteger
            Dim Hangar As UShort
            Dim State As UShort
            Dim Unused As UInteger
            Dim Link As UInteger
        End Structure

        <StructLayout(LayoutKind.Sequential)> Public Structure LocationData '20 Bytes
            Dim Left As UInteger
            Dim Top As UInteger
            Dim Right As UInteger
            Dim Bottom As UInteger
            Dim NameString As UShort
            Dim Flags As UShort
        End Structure

        <StructLayout(LayoutKind.Sequential)> Public Structure Trigger  '2400 Bytes
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=16)> _
            Public Conditions() As Condition
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=64)> _
            Public Actions() As Action
            Public Flags As UInteger
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=28)> _
            Public Players() As Byte

            Shared Function Create() As Trigger
                Dim Trig As New Trigger
                Trig.Conditions = Array.CreateInstance(GetType(Condition), 16)
                Trig.Actions = Array.CreateInstance(GetType(Action), 64)
                Trig.Players = Array.CreateInstance(GetType(Byte), 28)
                Return Trig
            End Function

            <StructLayout(LayoutKind.Sequential)> Public Structure Condition    '20 Bytes    
                Public Location As UInteger
                Public Group As UInteger
                Public Parameter As UInteger
                Public Unit As UShort
                Public Cmp As Byte
                Public Condition As ConditionData
                Public Type As Byte
                Public Flags As Byte
                Public LinkedList As UShort
            End Structure

            <StructLayout(LayoutKind.Sequential)> Public Structure Action       '32 Bytes    
                Public Location As UInteger
                Public MsgString As UInteger
                Public WavString As UInteger
                Public Time As UInteger
                Public Group As UInteger
                Public Parameter1 As UInteger
                Public Type As Short
                Public Action As TActionData ' MActionData
                Public Parameter2 As Byte
                Public Flags As Byte
                Public Unused As Byte
                Public LinkedList As UShort
            End Structure
        End Structure

        Public Enum ConditionData As Byte
            NoCondition = 0
            CountdownTimer = 1
            Command = 2
            Bring = 3
            Accumulate = 4
            Kill = 5
            CommandMost = 6
            CommandMostAt = 7
            MoskKills = 8
            HighestScore = 9
            MostResources = 10
            Switch = 11
            ElapsedTime = 12
            MissionBriefing = 13
            Opponents = 14
            Deaths = 15
            CommandLeast = 16
            CommandLeastAt = 17
            LeastKills = 18
            LowestScore = 19
            LeastResources = 20
            Score = 21
            Always = 22
            Never = 23
        End Enum

        Public Enum TActionData As Byte
            NoAction = 0
            Victory = 1
            Defeat = 2
            PreserveTrigger = 3
            Wait = 4
            PauseGame = 5
            UnpauseGame = 6
            Transmission = 7
            PlayWav = 8
            DisplayTextMessage = 9
            CenterView = 10
            CreateUnitwithProperties = 11
            SetMissionObjectives = 12
            SetSwitch = 13
            SetCountdownTimer = 14
            RunAIScript = 15
            RunAIScriptAtLocation = 16
            LeaderBoradControl = 17
            LeaderBoradControlAtLocation = 18
            LeaderBoradResources = 19
            LeaderBoradKills = 20
            LeaderBoradPoints = 21
            KillUnit = 22
            KillUnitAtLocation = 23
            RemoveUnit = 24
            RemoveUnitAtLocation = 25
            SetResources = 26
            SetScore = 27
            MinimapPing = 28
            TalkingPortrait = 29
            MuteUnitSpeech = 30
            UnmuteUnitSpeech = 31
            LeaderBoradComputerPlayers = 32
            LeaderBoradGoalControl = 33
            LeaderBoradGoalControlAtLocation = 34
            LeaderBoradGoalResources = 35
            LeaderBoradGoalKills = 36
            LeaderBoradGoalPoints = 37
            MoveLocation = 28
            MoveUnit = 29
            LeaderBoradGreed = 40
            SetNextScenario = 41
            SetDoodadState = 42
            SetInvincibility = 43
            CreateUnit = 44
            SetDeaths = 45
            Order = 46
            Comment = 47
            GiveUnit = 48
            ModifyUnitHitPoints = 49
            ModifyUnitEnergy = 50
            ModifyUnitShieldPoints = 51
            ModifyUnitResourceAmount = 52
            ModifyUnitHangerCount = 53
            PauseTimer = 54
            UnpauseTimer = 55
            Draw = 56
            SetAllianceStatus = 57
            DisableDebug = 58
            EnableDebug = 59
        End Enum

        Public Enum MActionData
            NoAction = 0
            Wait = 1
            PlayWav = 2
            TextMessage = 3
            MissionObjectives = 4
            ShowPortrait = 5
            HidePortrait = 6
            DisplaySpeakingPortrait = 7
            Transmission = 8
            SkipTutorialEnabled = 9
        End Enum

        Public Function LastStringIndex(StrList As Dictionary(Of Integer, String)) As Integer
            Dim Max As Integer
            For Each Index In StrList.Keys
                Max = Math.Max(Max, Index)
            Next
            Return Max
        End Function

        Public Function StructureCreate(Of StructType)(ByteData() As Byte) As StructType
            Dim Struct As StructType, Buffer As IntPtr, Size As Integer
            Size = Marshal.SizeOf(GetType(StructType))
            If ByteData.Length < Size Then Return Nothing
            Buffer = Marshal.AllocHGlobal(Size)
            Marshal.Copy(ByteData, 0, Buffer, Size)
            Struct = Marshal.PtrToStructure(Buffer, GetType(StructType))
            Marshal.FreeCoTaskMem(Buffer)
            Return Struct
        End Function

        Public Function StructureToByte(Of StructType)(Struct As StructType) As Byte()
            Dim ByteData() As Byte, Buffer As IntPtr, Size As Integer
            Size = Marshal.SizeOf(GetType(StructType))
            ByteData = Array.CreateInstance(GetType(Byte), Size)
            Buffer = Marshal.AllocHGlobal(Size)
            Marshal.StructureToPtr(Struct, Buffer, False)
            Marshal.Copy(Buffer, ByteData, 0, Size)
            Marshal.FreeCoTaskMem(Buffer)
            Return ByteData
        End Function

        Public Function StructureArrayCreate(Of StructType)(Section As CSection) As StructType()
            Dim StructSize As Integer, Size As Integer, i As Integer
            Dim StructArray() As StructType
            StructSize = Marshal.SizeOf(GetType(StructType))
            Size = Section.Data.Length / StructSize
            Section.Data.Position = 0
            StructArray = Array.CreateInstance(GetType(StructType), Size)
            For i = 0 To Size - 1
                StructArray(i) = StructureCreate(Of StructType)(Section.Reader.ReadBytes(StructSize))
            Next
            Return StructArray
        End Function

        Public Sub StructureArrayWrite(Of StructType)(Section As CSection, Struct() As StructType)
            Dim i As Integer
            Section.Data.SetLength(0)
            For i = 0 To Struct.Length - 1
                Section.Writer.Write(StructureToByte(Of StructType)(Struct(i)))
            Next
        End Sub
    End Module
End Namespace