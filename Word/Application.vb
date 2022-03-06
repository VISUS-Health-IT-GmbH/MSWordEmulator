<ComClass(Application.ClassId, Application.InterfaceId, Application.EventsId)>
Public Class Application
#Region "COM-GUIDs"
    ' Diese GUIDs stellen die COM-Identität für diese Klasse 
    ' und ihre COM-Schnittstellen bereit. Wenn Sie sie ändern, können vorhandene 
    ' Clients nicht mehr auf die Klasse zugreifen.
    Public Const ClassId As String = "50a42c0c-d50e-4fe0-b18e-39708934c7d7"
    Public Const InterfaceId As String = "e5795cf3-37d8-422c-be59-4fc587a1ddd9"
    Public Const EventsId As String = "efce4e3c-1990-4643-bb84-6e4f84193d97"
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
        Console.WriteLine(
            "[word.application.information] Emulation of 'word.application' COM object!"
        )
    End Sub


    ' Variables to be emulated
    ' ========================
    Public Visible As Boolean = False
    Public Documents As New Document
End Class
