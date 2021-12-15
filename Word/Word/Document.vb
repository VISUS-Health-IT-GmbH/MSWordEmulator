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
        Dim docObj As New Document()
        Return docObj
    End Function

    ' Emulation of "word.documents.saveas" subroutine
    ' TODO: Also emulate COMExceptions thrown by actual subroutine!
    Public Sub saveas(ByVal output As String, ByVal type As Integer)
        Console.WriteLine("Emulated 'word.documents.saveas' subroutine!")
    End Sub

    ' Emulation of "word.documents.close" subroutine
    Public Sub close()
        Console.WriteLine("Emulated 'word.documents.close' subroutine!")
    End Sub
End Class
