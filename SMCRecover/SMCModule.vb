Imports SMCRecover.ChunkModule

Module SMCModule
    Public Function Unprotect(ByRef PChunk As ChunkData) As Boolean
        Dim UChunk As New ChunkData
        Dim Param() As Object
        Param = New Object() {PChunk, UChunk}
        For Each Func In GetType(SMCModule).GetMethods(-1)
            If Func.Name.StartsWith("Chk_") Then _
                    Func.Invoke(Nothing, Param)
        Next
        PChunk = UChunk
        Return True
    End Function

    Private Sub Chk_CreateVersions(PChunk As ChunkData, UChunk As ChunkData)
        Dim Ver As UShort = PChunk.Version
        UChunk(UChunk.Add("TYPE")).Writer.Write(System.Text.Encoding.ASCII.GetBytes(IIf(Ver = &HCD, "RAWB", "RAWS")))
        UChunk(UChunk.Add("VER ")).Writer.Write(Ver)
        If Not Ver = &HCD Then UChunk(UChunk.Add("IVER")).Writer.Write(CUShort(IIf(Ver = &H2F Or Ver = &H39, 9, 10)))
        UChunk(UChunk.Add("IVE2")).Writer.Write(11US)
    End Sub

    Private Sub Chk_RecoverOpcode(PChunk As ChunkData, UChunk As ChunkData)
        UChunk.SectionList.Add(PChunk.GetOverwrite("VCOD"))
    End Sub

    Private Sub Chk_RecoverOwner(PChunk As ChunkData, UChunk As ChunkData)
        Dim Section As CSection
        Section = PChunk(PChunk.LastIndexOf("OWNR"))
        Section.Data.WriteTo(UChunk(UChunk.Add("IOWN")).Data)
        Section.Data.WriteTo(UChunk(UChunk.Add("OWNR")).Data)
    End Sub

    Private Sub Chk_RecoverERA(PChunk As ChunkData, UChunk As ChunkData)
        UChunk.Tileset = PChunk.Tileset
    End Sub

    Private Sub Chk_CreateTile(PChunk As ChunkData, UChunk As ChunkData)
        Dim Section As CSection
        Section = PChunk.GetOverwrite("MTXM")
        Section.Data.WriteTo(UChunk(UChunk.Add("TILE")).Data)
        Section.Data.WriteTo(UChunk(UChunk.Add("MTXM")).Data)
    End Sub

    Private Sub Chk_CreateISOM(PChunk As ChunkData, UChunk As ChunkData)
        Dim MSize As MapSize, Data As UShort
        Dim Section As CSection
        Section = UChunk(UChunk.Add("ISOM"))
        MSize = PChunk.MapSize
        Select Case PChunk.Tileset
            Case &H0, &H6, &H7
                Data = &H10
            Case &H1, &H2, &H3
                Data = &H20
            Case &H4, &H5
                Data = &H40
        End Select
        For i = 0 To (MSize.Width / 2 + 1) * (MSize.Height + 1) * 4 - 1
            Section.Writer.Write(Data)
        Next
    End Sub

    Private Sub Chk_RecoverSprite(PChunk As ChunkData, UChunk As ChunkData)
        Dim SpriteList As List(Of SpriteData)
        SpriteList = New List(Of SpriteData)(PChunk.Sprites)
        SpriteList.RemoveAll( _
            Function(Sprite As SpriteData) As Boolean
                Return Array.TrueForAll(Of Byte)(Starcraft.StructureToByte(Of SpriteData)(Sprite),
                    Function(Cmp As Byte) As Boolean
                        Return Cmp = &HFF
                    End Function)
            End Function)
        UChunk.Sprites = SpriteList.ToArray()
    End Sub

    Private Sub Chk_RecoverMask(PChunk As ChunkData, UChunk As ChunkData)
        Dim Section As CSection
        Dim Size As MapSize = PChunk.MapSize
        Dim i As Integer
        Section = UChunk(UChunk.Add("MASK"))
        PChunk(PChunk.IndexOf("MASK")).Data.WriteTo(Section.Data)
        For i = Section.Data.Position To Size.Height * Size.Width - 1
            Section.Data.WriteByte(&HFF)
        Next
        Section.Data.SetLength(Size.Height * Size.Width)
    End Sub

    Private Sub Chk_RecoverLocation(PChunk As ChunkData, UChunk As ChunkData)
        Dim Locations() As LocationData, i As Integer
        Locations = PChunk.Locations
        For i = 0 To Locations.Length - 1
            Locations(i).NameString = 0
        Next
        UChunk.Locations = Locations
    End Sub

    Private Sub Chk_RecoverTrigger(PChunk As ChunkData, UChunk As ChunkData)
        Dim Triggers As List(Of Trigger)
        Triggers = PChunk.Triggers
        Triggers.RemoveAll( _
            Function(Trig As Trigger)
                If Trig.Players(7) = 0 Then Return False
                If Not Trig.Conditions(0).Location = 64 Then Return False
                If Not Trig.Conditions(0).Group = 10 Then Return False
                If Not Trig.Conditions(0).Parameter = 3333 Then Return False
                If Not Trig.Conditions(0).Cmp = 10 Then Return False
                If Not Trig.Conditions(0).Condition = 3 Then Return False
                If Not Trig.Conditions(0).Flags = 16 Then Return False
                If Not Trig.Conditions(1).Condition = 0 Then Return False
                Return True
            End Function)
        UChunk.Triggers = Triggers
    End Sub

    Private Sub Chk_RecoverUnitName(PChunk As ChunkData, UChunk As ChunkData)
        Dim Data() As Byte, i As Integer, n As Integer
        Dim SName() As String = {"UNIS", "UNIx"}
        Dim Section As CSection
        For i = 0 To 1
            Section = PChunk(PChunk.IndexOf(SName(i)))
            If Section IsNot Nothing Then
                Data = Section.Data.ToArray()
                Section = UChunk(UChunk.Add(SName(i)))
                Section.Writer.Write(Data)
                Section.Data.Position = &HC78
                For n = 0 To 227
                    If Data(n) = 0 Then
                        Section.Writer.Write(0S)
                    Else
                        Section.Reader.ReadUInt16()
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub Chk_CopyData(PChunk As ChunkData, UChunk As ChunkData)
        Dim SName() As String = {"DIM ", "SIDE", "PUNI", "UPGR", "PTEC", "UNIT", "DD2 ", "UPRP", "UPUS", "MBRF", "SPRP", _
                                 "FORC", "WAV ", "UPGS", "TECS", "SWNM", "COLR", "PUPx", "PTEx", "UPGx", "TECx"}
        Dim Index As Integer
        For i = 0 To SName.Length - 1
            Index = PChunk.IndexOf(SName(i))
            If 0 <= Index Then UChunk.SectionList.Add(PChunk(Index))
        Next
    End Sub

    Private Sub Chk_CreateStrings(PChunk As ChunkData, UChunk As ChunkData)
        UChunk.SectionList.Add(PChunk.GetOverwrite("STR "))
        UChunk.SetStrings(UChunk.GetStrings())
    End Sub

    Private Sub Chk_SortSections(PChunk As ChunkData, UChunk As ChunkData)
        UChunk.Sort()
    End Sub
End Module
