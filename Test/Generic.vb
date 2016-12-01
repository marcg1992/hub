Imports System.IO
Imports System.Text
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization.Json
Imports System.Resources
Imports System.Globalization

<Serializable()> Public Enum CallOfDutyType As Byte
    CallOfDuty
    CallOfDutyUO 'Contains both CallOfDuty AND CallOfDutyUnitedOffensive
    CallOfDuty2
    ModernWarfare
    ModernWarfare2
    ModernWarfare3
    WorldAtWar
    BlackOps
    BlackOps2 'Not implemented yet
    BlackOps3 'Not implemented yet
    Ghosts    'Not implemented yet
    AdvancedWarfare 'Not implemented yet
    InfiniteWarfare 'Not implemented yet
End Enum

<Serializable()> Public Class CallOfDuty

    Private _name As String = "Call Of Duty"
    Private _type As CallOfDutyType = CallOfDutyType.CallOfDuty
    Private _supported As Boolean = True
    Private _regkey As String = "SOFTWARE\activision\Call of Duty"
    Private _regfolderkeyname As String = "installpath"
    Private _spexename As String = "CoDSP.exe"
    Private _mpexename As String = "CoDMP.exe"
    Private _archivefolders() As String = {"main"}
    Private _scriptfolders() As String = Nothing
    Private _archiveextension As String = ".pk3"
    Private _scriptcontainerextension As String = Nothing
    Private _textureextension As String = ".dds"
    Private _iwiversion As Byte = Nothing

    ReadOnly Property Name As String
        Get
            Return _name
        End Get
    End Property

    ReadOnly Property Folder As String
        Get
            Return DirectCast(My.Computer.Registry.LocalMachine.OpenSubKey(_regkey).GetValue(_regfolderkeyname), String)
        End Get
    End Property

    ReadOnly Property Supported As Boolean
        Get
            Return _supported
        End Get
    End Property

    ReadOnly Property RegistryKeyName As String
        Get
            Return _regkey
        End Get
    End Property

    ReadOnly Property RegistryFolderValueName As String
        Get
            Return _regfolderkeyname
        End Get
    End Property

    ReadOnly Property Type As CallOfDutyType
        Get
            Return _type
        End Get
    End Property

    ReadOnly Property SPExecutable As String
        Get
            Return _spexename
        End Get
    End Property

    ReadOnly Property MPExecutable As String
        Get
            Return _mpexename
        End Get
    End Property

    ReadOnly Property DataArchiveFolders As String()
        Get
            Return _archivefolders
        End Get
    End Property

    ReadOnly Property ScriptArchiveFolders As String()
        Get
            Return _scriptfolders
        End Get
    End Property

    ReadOnly Property DataArchiveExtension As String
        Get
            Return _archiveextension
        End Get
    End Property

    ReadOnly Property ScriptArchiveExtension As String
        Get
            Return _scriptcontainerextension
        End Get
    End Property

    ReadOnly Property IWITextureVersion As Byte
        Get
            Return _iwiversion
        End Get
    End Property

    ReadOnly Property ExistsOnPC As Boolean
        Get
            Try
                Dim CODFolder As String = Folder
                If File.Exists(CODFolder & "\" & _spexename) Or File.Exists(CODFolder & "\" & _mpexename) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property

    ReadOnly Property Installed As Boolean
        Get
            Try
                Dim CODFolder As String = Folder
                If Not CODFolder = Nothing Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property

    Friend Sub SaveJSON(FileName As String)
        Using JSONStream As New FileStream(FileName, FileMode.Create, FileAccess.Write)
            Dim JSONCreator As New DataContractJsonSerializer(Me.GetType)
            JSONCreator.WriteObject(JSONStream, Me)
        End Using
    End Sub

    Friend Sub SaveJSON(TargetStream As Stream, Optional CloseStream As Boolean = False)
        Dim JSONCreator As New DataContractJsonSerializer(Me.GetType)
        JSONCreator.WriteObject(TargetStream, Me)
    End Sub

    Friend Sub SaveBinary(FileName As String)
        Using BinaryStream As New FileStream(FileName, FileMode.Create, FileAccess.Write)
            Dim BinaryFormatter As New BinaryFormatter()
            BinaryFormatter.Serialize(BinaryStream, Me)
        End Using
    End Sub

    Friend Sub SaveBinary(TargetStream As Stream, Optional CloseStream As Boolean = False)
        Dim BinaryFormatter As New BinaryFormatter()
        BinaryFormatter.Serialize(TargetStream, Me)
        If CloseStream Then TargetStream.Close()
    End Sub

    Friend Shared Function FromJSON(SourceStream As Stream, Optional CloseStream As Boolean = False) As CallOfDuty
        Dim Output As CallOfDuty = DirectCast((New DataContractJsonSerializer(GetType(CallOfDuty))).ReadObject(SourceStream), CallOfDuty)
        If CloseStream Then SourceStream.Close()
        Return Output
    End Function

    Friend Shared Function FromJSON(FileName As String) As CallOfDuty
        Using JSONStream As New FileStream(FileName, FileMode.Open, FileAccess.Read)
            Return DirectCast((New DataContractJsonSerializer(GetType(CallOfDuty))).ReadObject(JSONStream), CallOfDuty)
        End Using
    End Function

    Friend Shared Function FromBinary(FileName As String) As CallOfDuty
        Using BinaryStream As New FileStream(FileName, FileMode.Open, FileAccess.Read)
            Return DirectCast((New BinaryFormatter().Deserialize(BinaryStream)), CallOfDuty)
        End Using
    End Function

    Friend Shared Function FromBinary(SourceStream As Stream, Optional CloseStream As Boolean = False) As CallOfDuty
        Dim Output As CallOfDuty = DirectCast(New BinaryFormatter().Deserialize(SourceStream), CallOfDuty)
        If CloseStream Then SourceStream.Close()
        Return Output
    End Function

    Shared Function GetAll() As CallOfDuty()
        Dim CallOfDutys As New List(Of CallOfDuty)
        Dim MyResourcesManager As ResourceManager = My.Resources.ResourceManager
        Dim MyResourceSet As ResourceSet = MyResourcesManager.GetResourceSet(CultureInfo.CurrentCulture, True, True)
        Dim MyResourceEntry As DictionaryEntry
        For Each MyResourceEntry In MyResourceSet
            If MyResourceEntry.Key.ToString.StartsWith("COD_") Then
                Try
                    Dim MyResourceBytes() As Byte = DirectCast(MyResourceEntry.Value, Byte())
                    Using MyResourceStream As New MemoryStream(MyResourceBytes)
                        CallOfDutys.Add(FromJSON(MyResourceStream, False))
                    End Using
                Catch ex As Exception
                    MsgBox("Could not load the following Call Of Duty .JSON file:" & vbNewLine & MyResourceEntry.Key.ToString & vbNewLine & ex.Message)
                End Try
            End If
        Next
        Return CallOfDutys.ToArray
    End Function

    Shared Function GetInstalled() As CallOfDuty()
        Dim Output As New List(Of CallOfDuty)
        For Each COD As CallOfDuty In GetAll()
            If COD.Installed Then
                Output.Add(COD)
                If COD.Type = CallOfDutyType.CallOfDutyUO Then
                    Output.Add(FromType(CallOfDutyType.CallOfDuty))
                End If
            End If
        Next
        Return Output.ToArray
    End Function

    Shared Function GetInstalledTypes() As CallOfDutyType()
        Dim Output As New List(Of CallOfDutyType)
        For Each COD As CallOfDuty In GetInstalled()
            Output.Add(COD.Type)
        Next
        Return Output.ToArray
    End Function

    Shared Function GetExisting() As CallOfDuty()
        Dim Output As New List(Of CallOfDuty)
        For Each COD As CallOfDuty In GetAll()
            If COD.ExistsOnPC Then
                Output.Add(COD)
                If COD.Type = CallOfDutyType.CallOfDutyUO Then
                    Output.Add(FromType(CallOfDutyType.CallOfDuty))
                End If
            End If
        Next
        Return Output.ToArray
    End Function

    Shared Function GetExistingTypes() As CallOfDutyType()
        Dim Output As New List(Of CallOfDutyType)
        For Each COD As CallOfDuty In GetExisting()
            Output.Add(COD.Type)
        Next
        Return Output.ToArray
    End Function

    Shared Function FromType(CODType As CallOfDutyType) As CallOfDuty
        Return FromJSON(New MemoryStream(DirectCast(My.Resources.ResourceManager.GetObject("COD_" & CODType.ToString), Byte())), True)
    End Function

    Private Sub New()

    End Sub

End Class

Friend Module Generic

    Public Function FormatLines(Input As String) As String()
        Dim Output As New List(Of String)
        Using InputStream As New MemoryStream(ASCIIEncoding.ASCII.GetBytes(Input))
            Using InputReader As New StreamReader(InputStream)
                While Not InputReader.EndOfStream
                    Output.Add(InputReader.ReadLine)
                End While
            End Using
        End Using
        Return Output.ToArray
    End Function

    Public Function FormatLines(Input() As String) As String
        Using OutputMemory As New MemoryStream()
            Using OutputWriter As New StreamWriter(OutputMemory)
                For Each InputLine As String In Input
                    OutputWriter.WriteLine(InputLine)
                Next
            End Using
            Return ASCIIEncoding.ASCII.GetString(OutputMemory.ToArray)
        End Using
    End Function

End Module