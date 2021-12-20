<ComClass(Document.ClassId, Document.InterfaceId, Document.EventsId)>
Public Class Document
#Region "COM-GUIDs"
    ' Diese GUIDs stellen die COM-Identität für diese Klasse 
    ' und ihre COM-Schnittstellen bereit. Wenn Sie sie ändern, können vorhandene 
    ' Clients nicht mehr auf die Klasse zugreifen.
    Public Const ClassId As String = "2605a61c-4ad9-4fa3-8f7c-ac5017e7a6ff"
    Public Const InterfaceId As String = "648eaed9-fb7b-4521-ba68-eb94df3d3ab8"
    Public Const EventsId As String = "5b316c80-2a78-40b4-aa0c-a61941789c3b"
#End Region


    ' Eine erstellbare COM-Klasse muss eine Public Sub New() 
    ' ohne Parameter aufweisen. Andernfalls wird die Klasse 
    ' nicht in der COM-Registrierung registriert und kann nicht 
    ' über CreateObject erstellt werden.
    Public Sub New()
        MyBase.New()
    End Sub


    ' (Necessary) public information subroutine
    Public Sub Information()
        Console.WriteLine("Emulation of 'word.document' COM object!")
    End Sub


    ' Functions and subroutines to be emulated
    ' ========================================
#Disable Warning IDE0060 ' Possible unused parameters
#Disable Warning IDE1006 ' Doesn't matter if functions / subroutines are upper case

    ' Emulation of "word.documents.open" function
    Public Function open(ByVal input As String) As Document
        Console.WriteLine("Emulated 'word.documents.open' function!")

        ' Check input for correctness
        If input Is Nothing Then
            ' 1) input NULL return NULL
            Return Nothing
        ElseIf Directory.Exists(input) Then
            ' 2) input directory throw COMException
            Throw New COMException
        ElseIf Not File.Exists(input) Then
            ' 3) input not a file throw COMException
            Throw New COMException
        End If

        ' Check if file is DOC / DOCX file by magic number
        ' DOC:  D0 CF 11 E0 A1 B1 1A E1
        ' DOCX: 50 4B 03 04  /  50 4B 05 06  /  50 4B 07 08
        Dim buffer() As Byte = New Byte(8) {}
        Using fs As New FileStream(input, FileMode.Open, FileAccess.Read, FileShare.None)
            fs.Read(buffer, 0, buffer.Length)
        End Using

        Dim sequences As New List(Of Byte())()
        sequences.Add(New Byte() {&HD0, &HCF, &H11, &HE0, &HA1, &HB1, &H1A, &HE1})
        sequences.Add(New Byte() {&H50, &H4B, &H3, &H4})
        sequences.Add(New Byte() {&H50, &H4B, &H5, &H6})
        sequences.Add(New Byte() {&H50, &H4B, &H7, &H8})

        Do
            For Each sequence As Byte() In sequences
                For i = 0 To UBound(sequence)
                    If sequence(i) <> buffer(i) Then
                        Dim message As String = "'word.documents.open' -> File '" + input + "' is no DOC / DOCX, see magic number: " + String.Join("", buffer)
                        Console.WriteLine(message)
                        Exit Do
                    End If
                Next
            Next
        Loop While False

        Dim docObj As New Document()
        Return docObj
    End Function


    ' Emulation of "word.documents.saveas" subroutine
    ' TODO: Check if output is a PDF file and overwrite if true!
    Public Sub saveas(ByVal output As String, ByVal type As Integer)
        Console.WriteLine("Emulated 'word.documents.saveas' subroutine!")

        ' Check input for correctness
        If output Is Nothing Then
            ' 1) input NULL throw COMException
            Throw New COMException
        ElseIf Directory.Exists(output) Then
            ' 2) input directory throw COMException
            Throw New COMException
        ElseIf File.Exists(output) Then
            ' 3) input an existing file throw COMException
            Throw New COMException
        End If

        ' Create pseude PDF file with magic number
        ' PDF:  25 50 44 46 2D
        Dim buffer() As Byte = New Byte() {&H25, &H50, &H44, &H46, &H2D}
        File.WriteAllBytes(output, buffer)
    End Sub


    ' Emulation of "word.documents.close" subroutine
    Public Sub close()
        Console.WriteLine("Emulated 'word.documents.close' subroutine!")
    End Sub
End Class
