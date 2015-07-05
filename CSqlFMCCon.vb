Imports System.Data.SqlClient
Imports System.data.OleDb





Public Class CSqlFMCCon
    Public errores As String
Sub cargaListViewSQL(ByVal dr As SqlDataReader, ByRef lv As ListView)
        Try
            lv.Clear()
            lv.View = View.Details
            For i As Integer = 0 To dr.FieldCount - 1
                lv.Columns.Add(dr.GetName(i), 70, HorizontalAlignment.Left)
            Next
            Dim k, j
            Do While dr.Read
                For j = 0 To dr.FieldCount - 1
                    If j = 0 Then
                        lv.Items.Add(dr(j))
                    Else
                        lv.Items(k).SubItems.Add(dr(j))
                    End If
                Next
                k += 1
            Loop
        Catch ex As Exception
            'MsgBox("cargaListView: " & ex.Message)
            errores = errores & "muestraConstrains: " & ex.Message
        End Try
     
    End Sub

    Sub cargaListViewOLE(ByVal dr As OleDbDataReader, ByRef lv As ListView)
        Try
            lv.Clear()
            lv.View = View.Details
            For i As Integer = 0 To dr.FieldCount - 1
                lv.Columns.Add(dr.GetName(i), 70, HorizontalAlignment.Left)
            Next
            Dim k, j
            Do While dr.Read
                For j = 0 To dr.FieldCount - 1
                    If j = 0 Then
                        lv.Items.Add(dr(j))
                    Else
                        lv.Items(k).SubItems.Add(dr(j))
                    End If
                Next
                k += 1
            Loop
        Catch ex As Exception
            'MsgBox("cargaListView: " & ex.Message)
            errores = errores & "muestraConstrains: " & ex.Message
        End Try

    End Sub

    Function ejecutaQuerySQL(ByVal cn As SqlConnection, ByVal query As String) As SqlDataReader
        'pre: cn abierto
        Try
            Dim nlineas As Integer
            Dim cmd As New SqlCommand(query, cn)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader()
            Return dr
        Catch ex As Exception
            'MsgBox("ejecutaQuerySQL: " & ex.Message)
            errores = errores & "ejecutaQuerySQL: " & ex.Message
        End Try

    End Function
    Function ejecutaQueryOledb(ByVal cn As OleDbConnection, ByVal query As String) As OleDbDataReader
        'pre: cn abierto

        Try
            'cn.Open()
            Dim cmd As New OleDbCommand(query, cn)
            ejecutaQueryOledb = cmd.ExecuteReader()
            'cn.Close()
        Catch ex As Exception
            'MsgBox("ejecutaQueryOledb: " & ex.Message)
            errores = errores & "ejecutaQueryOledb: " & ex.Message
        End Try

    End Function

     Function ejecutaQuerySQLScalar(ByVal cn As SqlConnection, ByVal query As String) As Integer
        'pre: cn abierto
        Try
            Dim result As Integer
            Dim cmd As New SqlCommand(query, cn)
            result = cmd.ExecuteScalar
            Return result
        Catch ex As Exception
            'MsgBox("ejecutaQuerySQL: " & ex.Message)
            errores = errores & "ejecutaQuerySQL: " & ex.Message
            Return -1
        End Try

    End Function

    Function ejecutaQueryOledbScalar(ByVal cn As OleDbConnection, ByVal query As String) As Integer
        'pre: cn abierto

        Try
            'cn.Open()
            Dim result As Integer
            Dim cmd As New OleDbCommand(query, cn)
            result  = cmd.ExecuteScalar
            Return result
            ' cn.Close()
        Catch ex As Exception
            'MsgBox("ejecutaQueryOledbScalar: " & ex.Message)
            errores = errores & "ejecutaQueryOledbScalar: " & ex.Message
            Return -1
        End Try

    End Function

    Function ejecutaQueryOledbnonExecute(ByVal cn As OleDbConnection, ByVal query As String) As Integer
        Try
            ' cn.Open()
            Dim cmd As New OleDbCommand(query, cn)
            ejecutaQueryOledbnonExecute = cmd.ExecuteNonQuery
            ' cn.Close()
        Catch ex As Exception
            'MsgBox("ejecutaQueryOledbnonExecute " & ex.Message)
            errores = errores & "ejecutaQueryOledbnonExecute: " & ex.Message
        End Try

    End Function
    Function ejecutaQuerySQLnonExecute(ByVal cn As SqlConnection, ByVal query As String) As Integer
        'pre: cn abierto
        Try
            'cn.Open()
            Dim cmd As New SqlCommand(query, cn)
            ejecutaQuerySQLnonExecute = cmd.ExecuteNonQuery
            'cn.Close()
        Catch ex As Exception
            'MsgBox("ejecutaQuerySQLnonExecute " & ex.Message)
            errores = errores & "ejecutaQuerySQLnonExecute: " & ex.Message
        End Try

    End Function



   Sub cargaTreeOledb(ByVal treevw As TreeView, ByVal cn As OleDbConnection, ByVal dr As OleDbDataReader)
        Try
            treevw.Nodes.Clear()
            Dim nodo As New TreeNode
            Do While dr.Read
                nodo = New TreeNode(dr(0))
                treevw.Nodes.Add(nodo)
            Loop
        Catch ex As Exception
            'MsgBox(" cargaTreeOledb: " & ex.Message)
            errores = errores & "cargaTreeOledb: " & ex.Message
        End Try


    End Sub


  Sub cargaTreeOledb(ByVal treevw As TreeView, ByVal cn As OleDbConnection, ByVal dr As OleDbDataReader, ByVal dr1 As OleDbDataReader)
        Try
            Dim cn1 As New OleDbConnection
            cn1.ConnectionString = cn.ConnectionString
            cn1.Open()
            cn.Open()


            treevw.Nodes.Clear()
            Dim nodo As New TreeNode
            Do While dr.Read
                nodo = New TreeNode(dr(0))
                treevw.Nodes.Add(nodo)
                Dim nodohijo As TreeNode
                Do While dr1.Read
                    nodohijo = New TreeNode(dr1(0))
                    nodo.Nodes.Add(nodohijo)
                Loop
                ' dr1(0)
            Loop

            cn.Close()
            cn1.Close()
        Catch ex As Exception
            MsgBox(" cargaTreeOledb2niveles: " & ex.Message)
            'errores = errores & "cargaTreeOledb2niveles: " & ex.Message
        End Try

    End Sub

    


    Function montaQuery(ByVal query As String, ByVal valor As String) As String
        query = query & "' = '" & valor & "'"
        Return query
    End Function

    Sub cargaTreeSql(ByVal treevw As TreeView, ByVal cn As SqlConnection, ByVal dr As SqlDataReader)
        Try
            treevw.Nodes.Clear()
            Dim nodo As New TreeNode
            Do While dr.Read
                nodo = New TreeNode(dr(0))
                treevw.Nodes.Add(nodo)
            Loop
        Catch ex As Exception
            'MsgBox(" cargaTreeSql: " & ex.Message)
            errores = errores & "cargaTreeSql: " & ex.Message
        End Try

    End Sub
    Sub CargarTablassql(ByVal cn As SqlConnection, ByVal treevw As TreeView)
        Try
            cn.Open()
            Dim query As String
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            query = "select * from information_schema.tables"
            cmd = New SqlCommand(query, cn)
            dr = cmd.ExecuteReader
            Do While dr.Read
                If dr("TABLE_TYPE") = "BASE TABLE" AndAlso dr("TABLE_NAME") <> "dtproperties" Then
                    treevw.Nodes.Add(dr("TABLE_NAME"))
                End If
            Loop
            cn.Close()
        Catch ex As Exception
            'MsgBox("CargarTablasAsql: " & ex.Message)
            errores = errores & "CargarTablassql: " & ex.Message
        End Try
     
    End Sub


    Sub CargarTablasAccess(ByVal cn As OleDbConnection, ByRef treevw As TreeView)
        Try
            cn.Open()
            Dim tabla As DataTable
            Dim fila As DataRow
            Dim col As DataColumn
            tabla = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

            For Each fila In tabla.Rows
                For Each col In tabla.Columns
                    If col.ColumnName = "TABLE_NAME" Then
                        treevw.Nodes.Add(fila(col))
                    End If
                Next
            Next
            cn.Close()
        Catch ex As Exception
            'MsgBox("CargarTablasAccess: " & ex.Message)
            errores = errores & "CargarTablasAccess: " & ex.Message
        End Try

    End Sub
    Function ejecutaqueryConectada(ByVal cn As OleDbConnection, ByVal query As String, ByVal tipo As Integer, ByRef numresult As Integer) As OleDbDataReader
        'pre: cn open
        'post:cn no se cierra,
        '       si ExecuteReader devuelve resultado y en nlineas el num decampos
        '       si escalar resultado en  numresult
        '       si executenonquery en numresult el num de lineas afectadas

        Try
            Dim dr As OleDbDataReader
            Dim cmd As New OleDbCommand(query, cn)
            Select Case tipo
                Case 0
                    dr = cmd.ExecuteReader()
                    numresult = dr.FieldCount
                    Return dr
                Case 1
                    numresult = cmd.ExecuteScalar
                    Return dr
                Case Else
                    numresult = cmd.ExecuteNonQuery()
                    Return dr
            End Select
        Catch ex As Exception

        End Try

    End Function

    'Sub GRABADATOS(ByVal TIPO As Integer)

    '    Select Case TIPO
    'AUTORES
    '   Case 1
    '          Dim i As Integer
    '          Dim f1 As New StreamWriter("Libros.txt", False) 'false machaca lo anterior
    '         For i = 0 To Me.arrAutores.Count - 1
    '            f1.WriteLine(CType(Me.arrAutores(i), CAutores).formatea)
    '       Next
    '      f1.Close()

    '  Case 2
    '     'CLIENTES
    '    Dim i As Integer
    '   Dim f1 As New StreamWriter("Clientes.txt", False) 'false machaca lo anterior
    '  For i = 0 To Me.arrclientes.Count - 1
    '       f1.WriteLine(CType(Me.arrclientes(i), CClientes).formatea(False))
    '    Next
    '     f1.Close()
    ' Case 3
    'alquiler
    '      Dim i As Integer
    '      Dim f1 As New StreamWriter("Alquiler.txt", False) 'false machaca lo anterior
    '      For i = 0 To Me.arrAlquileres.Count - 1
    '          f1.WriteLine(CType(Me.arrclientes(i), CAlquiLibro).formatea)
    '      Next
    '      f1.Close()
    '  Case Else
    '      'LIBROS'

    '        Dim i As Integer
    '        Dim f1 As New StreamWriter("Libros.txt", False) 'false machaca lo anterior
    '        For i = 0 To Me.arrAutores.Count - 1
    '            f1.WriteLine(CType(Me.arrLibros(i), Clibros).formatea(False))

    '        Next
    '        f1.Close()

    '    End Select

    ' Sub LEEDATOS(ByVal TIPO As Integer)
    '       Select Case TIPO
    '          'AUTORES
    '       Case 1
    '               If File.Exists("Autores.TXT") Then
    '                   Dim F1 As New StreamReader("Autores.TXT")
    '                   Dim linea As String = F1.ReadLine()
    '                   Dim o As CAutores
    '                   Dim v() As String
    '                   Do Until linea Is Nothing
    '                       o = New CAutores
    '                       v = linea.Split(sep)
    '                       o.nombre = v(0)
    '                       o.apellidos = v(1)
    '                       o.fechaNac = CDate(v(2))
    '                       Me.arrAutores.Add(o)
    '                       linea = F1.ReadLine()
    '                   Loop
    '                   F1.Close()
    '               End If
    '            Case 2
    '                If File.Exists("Clientes.TXT") Then
    '                    Dim F1 As New StreamReader("Clientes.TXT")
    '                    Dim linea As String = F1.ReadLine()
    '                    Dim o As CClientes
    '                    Dim v() As String
    '                    Do Until linea Is Nothing
    '                       o = New CClientes
    '                       v = linea.Split(sep)
    '                       o.nombre = v(0)
    '                       o.apellidos = v(1)
    '                       o.fechaing = CDate(v(2))
    '                       Me.arrclientes.Add(o)
    '                       linea = F1.ReadLine()
    '                   Loop
    '                   F1.Close()
    '               End If
    '           Case Else
    '               If File.Exists("Libros.TXT") Then
    '                    Dim F1 As New StreamReader("Libros.TXT")
    '                    Dim linea As String = F1.ReadLine()
    '                    Dim o As Clibros
    '                    Dim v() As String
    '                    Do Until linea Is Nothing
    '                        o = New Clibros
    '                        v = linea.Split(sep)
    '                        o.titulo = v(0)
    '                        o.autor = v(1)
    '                        o.precio = CDbl(v(2))
    '                        Me.arrLibros.Add(o)
    '                        linea = F1.ReadLine()
    '                    Loop
    '                    F1.Close()
    '                End If
    '        End Select


    'End Sub

End Class
