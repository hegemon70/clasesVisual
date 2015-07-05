Imports System.io
Public Class CErroresFMC
    'creado por Fernando Mendiz
    '29/11/2012 
    '12:09
    'v1.00
    'cambios : inicial
    'Para usar 
    'necesaria la clase GestorErroresFMC

   
    Public algoritmo As String
    Public archivo As String
    Public fechatxt As String
    Public descripcion As String
    Public data As Date


    Sub New(ByVal mensaje As String)
        Me.descripcion = "BAD_FORMAT:" & mensaje
        Me.algoritmo = Nothing
        Me.archivo = Nothing
        Me.fechatxt = Nothing
        Me.data = Nothing
    End Sub
    Sub New(ByVal algor As String, ByVal archiv As String, ByVal descrip As String, ByVal ftxt As String)
        Me.algoritmo = algor
        Me.archivo = archiv
        Me.fechatxt = ftxt
        Me.descripcion = descrip
        If IsDate(ftxt) Then
            Me.data = CDate(ftxt)
        Else
            Me.data = Nothing
        End If
    End Sub


    Function formatea(ByVal conSaltodeCarro As Boolean) As String

        Dim recog As String
        If Me.data = Nothing Then
            recog = fechatxt
        Else
            'recog = data.ToString("HH:mm:ss:dd:MM:yyyy")
            recog = data.ToString("yyyy:MM:dd:HH:mm:ss")
        End If
        recog = recog & sep & Me.algoritmo
        recog = recog & sep & Me.archivo
        recog = recog & sep & Me.descripcion
        If conSaltodeCarro Then
            recog = recog & vbCrLf
        End If
        Return recog
    End Function
End Class
