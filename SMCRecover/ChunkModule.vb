Imports System.IO

Namespace ChunkModule

    Partial Public Class ChunkData
        Implements IEnumerable(Of CSection)
        Public ReadOnly SectionList As List(Of CSection)

        Public Sub New()
            SectionList = New List(Of CSection)
        End Sub

        Public Sub New(ByteData() As Byte)
            SectionList = New List(Of CSection)
            ChunkData = ByteData
        End Sub

        Protected Overrides Sub Finalize()
            SectionList.Clear()
            MyBase.Finalize()
        End Sub

        Private Function ObjEnum() As IEnumerator Implements  IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function

        Public Function GetEnumerator() As IEnumerator(Of CSection) Implements IEnumerable(Of CSection).GetEnumerator
            Return SectionList.GetEnumerator()
        End Function

        Public Function toArray() As Byte()
            Return ChunkData
        End Function

        Public Function Add(Name As String) As Integer
            Dim Section As New CSection()
            Section.Name = Name
            SectionList.Add(Section)
            Return SectionList.Count - 1
        End Function

        Public Function Insert(Name As String, Index As Integer) As Integer
            Dim Section As New CSection()
            Section.Name = Name
            SectionList.Insert(Index, Section)
            Return Index
        End Function

        Public Function IndexOf(Name As String) As Integer
            Return SectionList.FindIndex( _
                Function(Section As CSection) As Boolean
                    Return Section.Name = Name
                End Function)
        End Function

        Public Function LastIndexOf(Name As String) As Integer
            Return SectionList.FindLastIndex( _
                Function(Section As CSection) As Boolean
                    Return Section.Name = Name
                End Function)
        End Function

        Public Sub Move(Source As Integer, Dest As Integer)
            Dim Section As CSection
            Section = SectionList(Source)
            SectionList.Insert(Dest, Section)
            SectionList.RemoveAt(Source)
        End Sub

        Public Sub Sort()
            SectionList.Sort()
        End Sub

        Public Sub Remove(Index As Integer)
            SectionList.RemoveAt(Index)
        End Sub

        Public Sub Remove(Name As String)
            SectionList.RemoveAll( _
                Function(Section As CSection) As Boolean
                    Return Section.Name = Name
                End Function)
        End Sub

        Public Function GetOverwrite(Name As String) As CSection
            Dim Section As New CSection
            Section.Name = Name
            For Each SectionData In Sections(Name)
                Section.Data.Position = 0
                SectionData.Data.WriteTo(Section.Data)
            Next
            Return Section
        End Function

        Public ReadOnly Property Count As Integer
            Get
                Return SectionList.Count
            End Get
        End Property

        Public Property ChunkData() As Byte()
            Get
                Dim Stream As MemoryStream, Writer As BinaryWriter
                Stream = New MemoryStream()
                Writer = New BinaryWriter(Stream)
                For Each Section In SectionList
                    Writer.Write(System.Text.Encoding.ASCII.GetBytes(Section.Name), 0, 4)
                    Writer.Write(CInt(Section.Data.Length))
                    Writer.Write(Section.Data.ToArray())
                    Writer.Flush()
                Next
                Return Stream.ToArray()
            End Get
            Set(ByteData As Byte())
                Dim Stream As MemoryStream, Reader As BinaryReader
                Stream = New MemoryStream(ByteData)
                Reader = New BinaryReader(Stream)
                SectionList.Clear()
                Do Until Stream.Length < Stream.Position + 8
                    Dim Section As New CSection, DSize As Integer
                    Section.Name = System.Text.Encoding.ASCII.GetString(Reader.ReadBytes(4))
                    DSize = Reader.ReadInt32()
                    If DSize > 0 Then
                        If Stream.Length > Stream.Position + DSize Then
                            Section.Writer.Write(Reader.ReadBytes(DSize))
                        Else
                            Section.Writer.Write(Reader.ReadBytes(Stream.Length - Stream.Position))
                        End If
                    Else
                        Stream.Seek(DSize, SeekOrigin.Current)
                    End If
                    SectionList.Add(Section)
                Loop
            End Set
        End Property

        Default Public Property Sections(Index As Integer) As CSection
            Get
                If Index < 0 OrElse SectionList.Count <= Index Then Return Nothing
                Return SectionList(Index)
            End Get
            Set(Section As CSection)
                SectionList(Index) = Section
            End Set
        End Property

        Default Public ReadOnly Property Sections(Name As String) As CSection()
            Get
                Return SectionList.FindAll( _
                    Function(Section As CSection) As Boolean
                        Return Section.Name = Name
                    End Function).ToArray()
            End Get
        End Property
    End Class

    Public Class CSection
        Implements IComparable(Of CSection)

        Public Name As String
        Public ReadOnly Data As MemoryStream
        Public ReadOnly Reader As BinaryReader
        Public ReadOnly Writer As BinaryWriter

        Public Sub New()
            Data = New MemoryStream()
            Reader = New BinaryReader(Data)
            Writer = New BinaryWriter(Data)
        End Sub

        Protected Overrides Sub Finalize()
            Data.Close()
            MyBase.Finalize()
        End Sub

        Public Overrides Function ToString() As String
            Return "Section """ & Name & """"
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Data.GetHashCode()
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean 'Don't use for speed
            Dim Pos As Integer, Eq As Boolean
            If Not obj.GetType() = Me.GetType() Then Return False
            If Not obj.Name = Name Then Return False
            If Not obj.Data.Length = Data.Length Then Return False
            Pos = Data.Position
            Data.Position = 0
            Eq = Array.TrueForAll(Of Byte)(obj.Data.toArray(),
                Function(Cmp As Byte) As Boolean
                    Return Cmp = Data.ReadByte()
                End Function)
            Data.Position = Pos
            Return Eq
        End Function

        Public Function CompareTo(op As CSection) As Integer Implements IComparable(Of CSection).CompareTo
            Dim Index1 As Integer, Index2 As Integer
            If Me.Name = op.Name Then Return 0
            Index1 = Array.IndexOf(Of String)(SectionSort, Me.Name)
            Index2 = Array.IndexOf(Of String)(SectionSort, op.Name)
            If Index1 < 0 And Index2 < 0 Then
                Return StrComp(Me.Name, op.Name)
            ElseIf Index1 < 0 Then
                Return 1
            ElseIf Index2 < 0 Then
                Return -1
            Else
                Return Index1.CompareTo(Index2)
            End If
        End Function
    End Class
End Namespace