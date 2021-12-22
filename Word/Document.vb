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
        Console.WriteLine("[word.document.information] Emulation of 'word.document' COM object!")
    End Sub


    ' Functions and subroutines to be emulated
    ' ========================================

    ' Emulation of "word.documents.open" function
    Public Function Open(ByVal input As String) As Document
        Console.WriteLine("[word.document.open] Emulated function!")

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
        Dim buffer() As Byte = New Byte(7) {}
        Using fs As New FileStream(input, FileMode.Open, FileAccess.Read, FileShare.None)
            fs.Read(buffer, 0, buffer.Length)
        End Using

        Dim sequences As New List(Of Byte()) From {
            New Byte() {&HD0, &HCF, &H11, &HE0, &HA1, &HB1, &H1A, &HE1},
            New Byte() {&H50, &H4B, &H3, &H4},
            New Byte() {&H50, &H4B, &H5, &H6},
            New Byte() {&H50, &H4B, &H7, &H8}
        }

        ' check DOCX magic number and all 3 possible DOC magic numbers
        Dim found As Boolean = CheckMagicNumbers(sequences(0), buffer) OrElse
                               CheckMagicNumbers(sequences(1), buffer) OrElse
                               CheckMagicNumbers(sequences(2), buffer) OrElse
                               CheckMagicNumbers(sequences(3), buffer)

        ' log warning if nothing matches (normal MS Word does not fail)
        If Not found Then
            Console.WriteLine(
                "[word.document.open - WARNING] File '" + input +
                "' is no DOC / DOCX, see magic number: " + MagicNumberToString(buffer)
            )
        End If

        Dim docObj As New Document()
        Return docObj
    End Function


    ' Emulation of "word.documents.saveas" subroutine
    ' TODO: Check if output is a PDF file and overwrite if true!
#Disable Warning IDE0060 ' Possible unused parameters
    Public Sub SaveAs(ByVal output As String, ByVal type As Integer)
        Console.WriteLine("[word.document.saveas] Emulated subroutine!")

        ' Check if file is PDF file by magic number
        ' PDF:  25 50 44 46 2D
        Dim buffer() As Byte = New Byte() {&H25, &H50, &H44, &H46, &H2D}

        ' Check input for correctness
        If output Is Nothing Then
            ' 1) input NULL throw COMException
            Throw New COMException
        ElseIf Directory.Exists(output) Then
            ' 2) input directory throw COMException
            Throw New COMException
        ElseIf File.Exists(output) Then
            ' 3) input an existing file throw COMException
            Dim input_buffer() As Byte = New Byte(4) {}
            Using fs As New FileStream(output, FileMode.Open, FileAccess.Read, FileShare.None)
                fs.Read(input_buffer, 0, input_buffer.Length)
            End Using

            If Not CheckMagicNumbers(buffer, input_buffer) Then
                Throw New COMException
            Else
                Console.WriteLine(
                    "[word.document.saveas - WARNING] File '" + output +
                    "' already exists as PDF, therefore will be overwritten!"
                )
            End If
        End If

        ' Write PDF magic number to file
        File.WriteAllBytes(output, buffer)
    End Sub
#Enable Warning IDE0060


    ' Emulation of "word.documents.close" subroutine
    Public Sub Close()
        Console.WriteLine("[word.document.close] Emulated subroutine!")
    End Sub


    ' Helper functions used in emulation
    ' ==================================

    ' Checks if two magic numbers match (saved as byte arrays of different length)
    Private Function CheckMagicNumbers(ByVal array As Byte(), ByVal reference As Byte()) As Boolean
        For i = 0 To UBound(array)
            If array(i) <> reference(i) Then
                Return False
            End If
        Next
        Return True
    End Function


    ' Converts a magic number into a string
    Private Function MagicNumberToString(ByVal array As Byte()) As String
        Dim builder As New StringBuilder(array.Length * 2)
        For Each b As Byte In array
            builder.Append(Conversion.Hex(b))
        Next
        Return builder.ToString()
    End Function
End Class
