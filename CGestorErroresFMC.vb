Imports System.io
Public Class GestorErroresFMC
    'creado por Fernando Mendiz
    '29/11/2012 
    '12:09
    'v1.00
    'cambios : inicial
    'Para usar 
    'necesaria la clase CErroresFMC
    'crear en un modulo 
    'public errores as string
    'declara un variable tipo GestorErroresFMC y ve colocando 
    'la funcion recogeErrores con 
    'este formato en los try catch
    ' cada mensaje de error deberia encabezarse por simbolo  #
    'y los siguentes campos separados por pipe
    'nombre de funcion o sub que produce el error 
    'nombre del archivo de esa funcion
    'descripcion del mensaje
    'now().toString
    'ejemplo:    "#recogeErrores|modulo|fallo desconocido|"& now().toString 
    Private Const sep As String = "|"
    Private Const sepM As String = "#"
    Public nomFichLog = "log.txt"

    Public arrlog As New ArrayList

    Sub New()

    End Sub
    Public Sub recogeErrores(ByVal message As String)

        'pre: cada mensaje de error deberia encabezarse por simbolo  #
        'y los siguentes campos separados por pipe
        'now().toString
        'nombre de funcion que produce el error 
        'nombre del archivo de esa funcion
        'descripcion del mensaje
        'ejemplo:    "#recogeErrores|modulo|fallo desconocido|"& now().toString 

        If errores <> Nothing Then
            errores = errores & message
        Else
            errores = message
        End If

    End Sub

    Public Sub rellenalog(ByVal errores As String)
        Try
            Dim v() As String
            Dim v1() As String
            Dim tam As Integer
            'Dim aviso As New CErroresFMC
            If errores <> Nothing Then
                v = errores.Split(sepM)
                tam = v.Length()
                For Each cursor As String In v
                    If cursor <> "" Then ' el primer elem de v esta vacio
                        Me.procesaMessage(cursor)
                    End If

                Next
            End If

        Catch ex As Exception
            MsgBox("#rellenaLog|GestorErroresFMC|" & ex.Message & "|" & Date.Now())
        End Try

    End Sub

    Public Sub creaNuevoLog()
        Try
            Me.reseteaOCreaFicheroSecuencial(Me.nomFichLog)
            Me.rellenalog(errores)
            Me.grabadatos(Me.arrlog, Me.nomFichLog, 1)
        Catch ex As Exception
            MsgBox("#creaNuevoLog||GestorErroresFMC|creando archivo log.txt:" & errores & "|" & Date.Now())
        End Try

    End Sub
    Function reseteaOCreaFicheroSecuencial(ByVal nomFich As String) As Integer
        Try
            If File.Exists(nomFich) Then
                Kill(nomFich)
            End If
            Return Me.creaFichero(nomFich)
        Catch ex As Exception
            MsgBox("#reseteaOCreaFicheroSecuencial|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())

        End Try

    End Function

    Sub leeLog()
        Try
            Me.leeDatos(Me.nomFichLog, sep, arrlog, 1)
        Catch ex As Exception
            MsgBox("#leeLog|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())
        End Try



    End Sub

    Function dameUltimoerror() As String
        Dim v() As String
        Dim recog As String = Nothing
        If errores <> Nothing Then
            v = errores.Split(sepM)
            If v.Length > 1 Then
                recog = v(v.Length - 1)
            Else
                recog = v(0)
            End If
        End If
        Return recog
        ' MsgBox("ultimoError:" & recog)
    End Function


    Sub muestraUltimoerror()
        Dim v() As String
        Dim recog As String
        If errores = Nothing Then
            MsgBox("no hay errores")
        Else
            v = errores.Split(sepM)
            If v.Length > 1 Then
                recog = v(v.Length - 1)
            Else
                recog = errores
            End If
            MsgBox("ultimoError:" & recog)
        End If


    End Sub

    Function creaFichero(ByVal nomFich As String) As Integer
        Try
            Dim numfich As Integer
            numfich = FreeFile()
            FileOpen(numfich, nomFich, OpenMode.Output, OpenAccess.Default)
            FileClose(numfich)
            Return numfich
        Catch ex As Exception
            MsgBox("#CreaFichero|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())
            Return -1
        End Try

    End Function

    Sub procesaMessage(ByVal message As String)
        Try
            'Dim aviso As New CErroresFMC
            Dim v1() As String
            v1 = message.Split(sep)
            If v1.Length = 4 Then
                Dim aviso As New CErroresFMC(v1(0), v1(1), v1(2), v1(3)) 'crea un aviso
                Me.arrlog.Add(aviso)
            Else
                Dim aviso As New CErroresFMC(message)
                Me.arrlog.Add(aviso)
            End If
        Catch ex As Exception
            MsgBox("#ProcesaMessage|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())
        End Try
    End Sub






    Function abreFicheroSecuencial(ByVal nombreFichHab As String) As Integer
        Dim filenumber As Integer
        Dim habAct As Habitacion
        Try
            filenumber = FreeFile()
            'tamregistro = Len(habAct)
            FileOpen(filenumber, nombreFichHab, OpenMode.Input, OpenAccess.ReadWrite, , )

            'tamFichero = FileLen(nombreFichHab)

        Catch ex As Exception
            'errores = ex.Message
            Me.recogeErrores("#abreFicheroSecuencial|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())
            Return -1
        End Try


        Return filenumber
    End Function

    Sub recogelineaEnArrayc1(ByVal v() As String, ByVal arr As ArrayList)
        Try
            Dim o As New CErroresFMC(v(0), v(1), v(2), v(3))
            'o.algoritmo = v(0)
            ' o.archivo = v(1)
            ' o.descripcion = v(2)
            ' o.fechatxt = v(3)
        Catch ex As Exception

            MsgBox("#recogelineaEnArrayc1|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())
        End Try

    End Sub

    Sub recogelineaEnArrayc2(ByVal v() As String, ByVal arr As ArrayList)
        Dim o As New CDetalleFac
        Try
            'o.precioHabSencilla = CDbl(v(0))
            'o.precioHabDoble = CDbl(v(1))
            'o.precioHabTriple = CDbl(v(2))
            'o.precioPensionCompleta = CDbl(v(3))
            'o.precioPensionMedia = CDbl(v(4))
            'o.precioPensionDesayuno = CDbl(v(5))
            'o.precioSupletoria = CDbl(v(6))
            'o.precioLavanderia = CDbl(v(7))
            'o.PrecioNevera = CDbl(v(8))
            'o.PrecioSpa = CDbl(v(9))
            arr.Add(o)
        Catch ex As Exception
            MsgBox("#recogelineaEnArrayc2|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())
        End Try

    End Sub



    Sub recogelineaEnArrayc3(ByVal v() As String, ByVal arr As ArrayList)
        'si fechas date

        'Dim o As New CErroresFMC



        ' n = 3
        'recog = recog & v(n) & sepFecha & v(n + 1) & sepFecha & (n + 2)
        ' o.fecha = CDate(recog)

        'arr.Add(o)
    End Sub
    Function dameValorBooleano(ByVal valor As Integer) As Boolean
        If valor = 0 Then
            Return False
        Else
            Return True
        End If
    End Function
    Function ponValorBooleano(ByVal valorBoo) As Integer
        If valorBoo Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Function damefecha(ByVal recog As String) As Date
        Return DateValue(recog)
    End Function
    Sub leeDatos(ByVal nombreArchivo As String, ByVal sep As String, ByVal arrayAnterior As ArrayList, ByVal tipo As Integer)
        Try
            If File.Exists(nombreArchivo) Then
                Dim f1 As New StreamReader(nombreArchivo)
                Dim linea As String
                linea = f1.ReadLine
                Dim v() As String

                Do Until linea = "" Or linea = Nothing
                    v = linea.Split(sep)
                    Select Case tipo
                        Case 1
                            recogelineaEnArrayc1(v, arrayAnterior)
                        Case 2
                            recogelineaEnArrayc2(v, arrayAnterior)
                        Case 3
                            recogelineaEnArrayc3(v, arrayAnterior)
                    End Select
                    linea = f1.ReadLine
                Loop
                f1.Close()
            End If
        Catch ex As Exception
            MsgBox("#leedatos|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())

            'errores = ex.Message

        End Try

    End Sub
    Sub grabadatos(ByVal array As ArrayList, ByVal nombreArchivo As String, ByVal tipo As Integer)
        Try
            Dim i As Integer
            Dim f1 As New StreamWriter(nombreArchivo, False) 'false machaca lo anterior
            For i = 0 To array.Count - 1
                Select Case tipo
                    Case Else
                        f1.WriteLine(CType(array(i), CErroresFMC).formatea(False))
                        ' Case 2
                        'f1.WriteLine(CType(array(i), CDetalleFac).formateaPrecios(0))
                        'Case 3
                        'f1.WriteLine(CType(array(i), CDetalleFac).formatea(False))
                End Select
            Next
            f1.Close()

        Catch ex As Exception
            MsgBox("#grabadatos|CGestorErroresFMC|" & ex.Message & "|" & Date.Now())
            'errores = ex.Message
        End Try
    End Sub
End Class

