'creado por Fernando Mendiz
'11/01/20123
'20:10
'v1.10


'Para usar hay que
'poner en primeralinea form Imports System.data.sqlclient
'buscar el objeto SqlConnection1 y ponerlo en el form
'Y DESPUEs de Inherits System.Windows.Forms.Form
'    Public cn As New SqlConnection
'enel load
'   cn.ConnectionString = Me.SqlConnection1.ConnectionString
'modificar la propiedad connectionstring  enlazandola con la basew de datos sqlserver

'cambios v1.07: 'ejecutaquery con ds como parametro
'cambios v1.08: 'cargaListView
'cambios v1.09: CargarTablasAccess tablassql cargaListViewOLE
'cambios v1.10: dameSqlcampo, dameSqlRegistro y dameNombreCampoformalSQL
Imports System.data.sqlclient
Imports System.data.OleDb

Public Class CSQLfmc
    Public errores As String

    Function ejecutaquery(ByVal cn As SqlConnection, ByVal query As String) As SqlDataReader
        Try
            Dim dr As SqlDataReader
            Dim cmd As New SqlCommand(query, cn)
            dr = cmd.ExecuteReader()
            Return dr
        Catch ex As Exception
            MsgBox("ejecutaquery: " & ex.Message)
        End Try


    End Function

    Function ejecutaquery(ByVal cn As OleDbConnection, ByVal query As String) As OleDbDataReader
        Try
            Dim dr As OleDbDataReader
            Dim cmd As New OleDbCommand(query, cn)
            dr = cmd.ExecuteReader()
            Return dr
        Catch ex As Exception
            MsgBox("ejecutaquery: " & ex.Message)
        End Try


    End Function
    Function ejecutaQuery(ByVal cn As SqlConnection, ByRef ds As DataSet, ByVal query As String, ByRef tipo As Integer, ByVal nombreTabla As String) As Integer
        Try
            Dim nlineas As Integer
            Dim da As New SqlDataAdapter(query, cn)
            nlineas = da.Fill(ds, nombreTabla)
            nlineas = da.Update(ds)
            Return nlineas
        Catch ex As Exception
            'MsgBox("ejecutaQuery: " & ex.Message)
            If ex.Message = "Update no puede encontrar TableMapping['Table'] o DataTable 'Table'." Then
                errores = Nothing
            Else
                errores = errores & "ejecutaQuery: " & ex.Message
            End If

        End Try

    End Function

    Function ejecutaQuery(ByVal cn As SqlConnection, ByRef ds As DataSet, ByVal query As String, ByVal nombreTabla As String) As Integer
        Try
            Dim nlineas As Integer
            Dim da As New SqlDataAdapter(query, cn)
            nlineas = da.Fill(ds, nombreTabla)
            Return nlineas
        Catch ex As Exception
            'MsgBox("ejecutaQuery: " & ex.Message)
            errores = errores & "ejecutaQuery: " & ex.Message
        End Try

    End Function

    Function ejecutaQuery(ByVal cn As SqlConnection, ByRef ds As DataSet, ByVal query As String, ByVal nombreTabla As String, ByVal nomp1 As String, ByVal tipop1 As String, ByVal valuep1 As String) As Integer
        Try
            Dim nlineas As Integer
            Dim cmd As New SqlCommand(query, cn)
            Dim da As New SqlDataAdapter(cmd)
            cmd.Parameters.Add(nomp1, tipop1)
            cmd.Parameters(0).Value = valuep1
            nlineas = da.Fill(ds, nombreTabla)
            Return nlineas
        Catch ex As Exception
            'MsgBox("ejecutaQuery1Param: " & ex.Message)
            errores = errores & "ejecutaQuery1Param: " & ex.Message
        End Try

    End Function
    Function ejecutaQuery(ByVal cn As SqlConnection, ByRef ds As DataSet, ByVal query As String, ByVal nombreTabla As String, ByVal nomp1 As String, ByVal tipop1 As String, ByVal valuep1 As String, ByVal nomp2 As String, ByVal tipop2 As String, ByVal valuep2 As String) As Integer
        Try
            Dim nlineas As Integer
            Dim cmd As New SqlCommand(query, cn)
            Dim da As New SqlDataAdapter(cmd)
            cmd.Parameters.Add(nomp1, tipop1)
            cmd.Parameters(0).Value = valuep1
            cmd.Parameters.Add(nomp2, tipop2)
            cmd.Parameters(1).Value = valuep2
            nlineas = da.Fill(ds, nombreTabla)
            Return nlineas
        Catch ex As Exception
            'MsgBox("ejecutaquery2param: " & ex.Message)
            errores = errores & "ejecutaquery2param: " & ex.Message
        End Try

    End Function
    Function ejecutaquery(ByVal cn As SqlConnection, ByRef ds As DataSet, ByVal query As String, ByVal nombretabla As String, ByVal posini As Integer, ByVal nreg As Integer) As Integer
        Try
            Dim nlineas As Integer
            Dim da As New SqlDataAdapter(query, cn)
            nlineas = da.Fill(ds, posini, nreg, nombretabla)
            Return nlineas
        Catch ex As Exception
            'MsgBox("ejecutaquery: " & ex.Message)
            errores = errores & "ejecutaQuery: " & ex.Message
        End Try

    End Function
    Function ejecutaquery(ByVal cn As SqlConnection, ByRef ds As DataSet, ByVal query As String, ByVal nombresmultipleT As String, ByVal numconsult As Integer) As Boolean
        'pre: en nombresmultiple tantos nombre consulta separados por , que ; en la query
        Try
            Dim v() As String
            v = nombresmultipleT.Split(",")
            If numconsult <> v.Length Then
                MsgBox("numero consulta no coincide con nombres mutiples")
                Return False
            Else
                Dim da As New SqlDataAdapter(query, cn)
                For i As Integer = 0 To v.Length - 1
                    If i = 0 Then
                        da.Fill(ds, v(0))
                    Else
                        ds.Tables(i).TableName = v(i)
                    End If
                Next
                Return True
            End If

        Catch ex As Exception
            'MsgBox("ejecutaqueryMultiple: " & ex.Message)
            errores = errores & "ejecutaqueryMultiple: " & ex.Message
            Return False
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
    Function contenidoCelda(ByVal dgrid As DataGrid) As String
        Try
            Dim celda As DataGridCell
            celda = dgrid.CurrentCell
            Return dgrid(celda.RowNumber(), celda.ColumnNumber())
        Catch ex As Exception
            'MsgBox("contenidoCelda: " & ex.Message)
            errores = errores & "contenidoCelda: " & ex.Message
        End Try

    End Function

    Function contenidoCelda(ByVal dgrid As DataGrid, ByVal numcolumna As Integer) As String
        Try
            Dim celda As DataGridCell
            celda = dgrid.CurrentCell
            Return dgrid(celda.RowNumber(), numcolumna)
        Catch ex As Exception
            'MsgBox("contenidoCeldacolum: " & ex.Message)
            errores = errores & "contenidoCeldacolum: " & ex.Message
        End Try

    End Function

    Function existeTabla(ByVal ds As DataSet, ByVal nombredatatable As String) As Boolean
        Try
            For Each tabla As DataTable In ds.Tables
                If tabla.TableName = nombredatatable Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox("existeTabla" & ex.Message)
            'errores = errores & "existeTabla:" & ex.Message
        End Try

    End Function

    Function creaColumnaEnTabla(ByRef tabla As DataTable, ByVal nombrecolumna As String, ByVal tip As SqlDbType) As Boolean
        Try
            Dim neocolumna As New DataColumn(nombrecolumna, tip.GetType())
            tabla.Columns.Add(neocolumna)
        Catch ex As Exception
            MsgBox("creaColumnaEnTabla" & ex.Message)
            'errores = errores & "creaColumnaEnTabla:" & ex.Message
        End Try

    End Function
    Function creaColumnaEnTabla(ByRef tabla As DataTable, ByVal nombrecolumna As String, ByVal tip As SqlDbType, ByVal tam As Integer) As Boolean
        Try

            Select Case tip
                Case SqlDbType.NVarChar
                    Dim neocolumna As New DataColumn(nombrecolumna, GetType(String))
                    neocolumna.MaxLength = tam
                    tabla.Columns.Add(neocolumna)
                Case SqlDbType.VarChar
                    Dim neocolumna As New DataColumn(nombrecolumna, GetType(String))
                    neocolumna.MaxLength = tam
                    tabla.Columns.Add(neocolumna)
                Case Else
                    Dim neocolumna As New DataColumn(nombrecolumna, tip.GetType)
                    tabla.Columns.Add(neocolumna)
            End Select
            Return True
        Catch ex As Exception

            'MsgBox("creaColumnaEnTablaTam" & ex.Message)
            errores = errores & "creaColumnaEnTablaTam:" & ex.Message
            Return False
        End Try

    End Function

    Function creaColumnaEnTabla(ByRef tabla As DataTable, ByVal nombrecolumna As String, ByVal tip As SqlDbType, ByVal esAutoincremental As Boolean) As Boolean
        Try
            Dim neocolumna As New DataColumn(nombrecolumna, tip.GetType())
            tabla.Columns.Add(neocolumna)
            If esAutoincremental Then
                neocolumna.AutoIncrement = True
                neocolumna.AutoIncrementSeed = 1
                neocolumna.AutoIncrementStep = 1
            End If
        Catch ex As Exception
            'MsgBox("creaColumnaEnTabla-autoincremental" & ex.Message)
            errores = errores & "creaColumnaEnTabla-autoincremental:" & ex.Message
        End Try

    End Function

    Function creaColumnaEnTabla(ByRef tabla As DataTable, ByVal nombrecolumna As String, ByVal tip As SqlDbType, ByVal esAutoincremental As Boolean, ByVal esPK As Boolean) As Boolean
        Try
            Dim neocolumna As New DataColumn(nombrecolumna, tip.GetType())
            tabla.Columns.Add(neocolumna)
            If esAutoincremental Then
                neocolumna.AutoIncrement = True
                neocolumna.AutoIncrementSeed = 1
                neocolumna.AutoIncrementStep = 1
            End If
            If esPK Then
                Dim CLAVE(1) As DataColumn
                CLAVE(0) = neocolumna
                tabla.PrimaryKey = CLAVE
            End If

        Catch ex As Exception
            'MsgBox("creaColumnaEnTabla-autoincremental-PK" & ex.Message)
            errores = errores & "creaColumnaEnTabla-autoincremental-PK:" & ex.Message
        End Try

    End Function

    Sub borraTablads(ByVal nombretabla As String, ByRef ds As DataSet)
        Try
            For Each tabla As DataTable In ds.Tables
                If tabla.TableName = nombretabla Then
                    tabla.Clear()
                End If
            Next
        Catch ex As Exception
            'MsgBox("borraTablads: " & ex.Message)
            errores = errores & "borraTablads: " & ex.Message
        End Try

    End Sub
    Sub creaRelacionesTabla(ByVal ds As DataSet, ByVal nombreRelacion As String, ByVal tablapadre As DataTable, ByVal tablahija As DataTable, ByVal nomCampoTpadre As String, ByVal nomCampoThija As String)
        Try
            Dim relacion As DataRelation = ds.Relations.Add(nombreRelacion, tablapadre.Columns(nomCampoTpadre), tablahija.Columns(nomCampoThija))

        Catch ex As Exception
            'MsgBox("creaRelacionesTabla: " & ex.Message)
            errores = errores & "creaRelacionesTabla: " & ex.Message
        End Try

        'ds.Relations.Add("nomRelacion",ds.tables("nomtabla").columns("nomcampo"),
    End Sub

    Function muestraConstraints(ByVal tabla As DataTable) As ArrayList
        Dim ar As New ArrayList
        For Each CLAVE As Constraint In tabla.Constraints
            ar.Add(CLAVE.ConstraintName)
        Next
        Return ar
    End Function

    Function selectdataview(ByVal ds As DataSet, ByVal nombretabla As String, ByVal condicion As String, ByVal valor As String, ByRef consulta As DataRow(), ByVal tipocampo As SqlDbType) As Boolean
        'pre: condicion= campo por el que buscar, valor valor concreto que tiene que tomar ese campo
        ' Dim consulta() As DataRow
        Try
            Select Case tipocampo
                Case SqlDbType.NVarChar
                    consulta = ds.Tables(nombretabla).Select(condicion & "'" & valor & "'")
                Case SqlDbType.Int
                    consulta = ds.Tables(nombretabla).Select(condicion & valor)
                Case SqlDbType.DateTime
                    consulta = ds.Tables(nombretabla).Select(condicion & "'" & valor & "'")
            End Select
            If consulta.Length = 0 Then

                'MsgBox("no encuentra nada por " & condicion & valor)
                errores = errores & "no encuentra nada por " & condicion & valor
                Return False
            Else
                Return True

            End If
        Catch ex As Exception
            errores = errores & "selectdataview: " & ex.Message
            Return False

        End Try

    End Function

    Function findRowsFMC(ByVal ds As DataSet, ByVal nombretabla As String, ByVal campoBusq As String, ByVal valorbusq As String, ByRef resultado() As DataRowView) As Integer
        Try
            Dim dv As DataView = New DataView(ds.Tables(nombretabla), "", campoBusq, DataViewRowState.CurrentRows)
            resultado = dv.FindRows(New Object() {valorbusq})
            Return resultado.Length

        Catch ex As Exception
            'MsgBox("findRowsFMC:" & ex.Message)
            errores = errores & "findRowsFMC:" & ex.Message
        End Try
    End Function


    Function CreaRelacion(ByRef ds As DataSet, ByVal nombreRelacion As String, ByVal TablaPral As String, ByVal TablaSecun As String, ByVal CampoPral As String, ByVal campoSecun As String)
        Try
            Dim rel As DataRelation = ds.Relations.Add(nombreRelacion, ds.Tables(TablaPral).Columns(CampoPral), ds.Tables(TablaSecun).Columns(campoSecun))
        Catch ex As Exception
            'MsgBox("CreaRelacion" & ex.Message)
            errores = errores & "CreaRelacion:" & ex.Message
        End Try

    End Function

    Function muestraRelacionEnListBox(ByVal ds As DataSet, ByRef lv As ListBox, ByVal nombreRelacion As String, ByVal RefTablaPral As String, ByVal campoTablaPral As String, ByVal campoTablaSecun As String)
        Try
            For Each drPral As DataRow In ds.Tables(RefTablaPral).Rows
                lv.Items.Add(drPral(campoTablaPral))
                For Each drSecun As DataRow In drPral.GetChildRows(nombreRelacion)
                    lv.Items.Add("-" & drSecun(campoTablaSecun))
                Next
            Next
        Catch ex As Exception
            'MsgBox("muestraRelacionEnListBox" & ex.Message)
            errores = errores & "muestraRelacionEnListBox:" & ex.Message
        End Try

    End Function

    Sub cargaListView(ByVal dr As SqlDataReader, ByRef lv As ListView)
        lv.Clear()
        lv.View = View.Details
        For i As Integer = 0 To dr.FieldCount - 1
            lv.Columns.Add(dr.GetName(i), 70, HorizontalAlignment.Left)
        Next
        Dim k, j As Integer

        Do While dr.Read
            For j = 0 To dr.FieldCount - 1
                If j = 0 Then
                    ' If IsDate(dr(j)) Then
                    lv.Items.Add(CStr(dr(j)))
                    ' Else'
                    'lv.Items.Add(dr(j))
                    ' End If



                Else
                    ' If IsDate(dr(j)) Then
                    lv.Items(k).SubItems.Add(CStr(dr(j)))
                    'Else


                End If

                ' End If
            Next
            k += 1
        Loop
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



    Sub Tablassql(ByVal cn As SqlConnection, ByVal tv As TreeView)
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
                    tv.Nodes.Add(dr("TABLE_NAME"))
                End If
            Loop
            cn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Sub Tablassql(ByVal cn As SqlConnection, ByVal L As ListBox)
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
                    L.Items.Add(dr("TABLE_NAME"))
                End If
            Loop
            cn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Sub Tablassql(ByVal cn As SqlConnection, ByVal Lv As ListView)
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
                    Lv.Items.Add(dr("TABLE_NAME"))
                End If
            Loop
            cn.Close()
        Catch ex As Exception

        End Try
    End Sub


    Sub CargarTablasAccess(ByVal cn As OleDbConnection, ByVal tv As TreeView)
        Try
            cn.Open()
            Dim tabla As DataTable
            Dim fila As DataRow
            Dim col As DataColumn
            tabla = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

            For Each fila In tabla.Rows
                For Each col In tabla.Columns
                    If col.ColumnName = "TABLE_NAME" Then
                        tv.Nodes.Add(fila(col))
                    End If
                Next
            Next
            cn.Close()
        Catch ex As Exception
            MsgBox("CargarTablasAccess: " & ex.Message)
        End Try

    End Sub

    Function SQLTABLASCAMPOSREG(ByVal CN As SqlConnection, ByVal tv As TreeView)
        'pre: debe existir las funciones dameSqlcampo, dameSqlRegistro y dameNombreCampoformalSQL
        'post:extrae todas las tablas de la base de datos 
        'todos los campos de cada tabla y todos los registros de casa campo
        ' y los mete en un treeview
        Try

            Dim cnnew As New SqlConnection
            Dim nodoPadre As TreeNode
            Dim query As String
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            cnnew.ConnectionString = CN.ConnectionString
            CN.Open()
            query = "select * from information_schema.tables"
            cmd = New SqlCommand(query, CN)
            dr = cmd.ExecuteReader
            Do While dr.Read
                If dr("TABLE_TYPE") = "BASE TABLE" AndAlso dr("TABLE_NAME") <> "dtproperties" Then
                    nodoPadre = New TreeNode(dr("TABLE_NAME"))

                    dameSqlcampo(cnnew, dr("TABLE_NAME"), nodoPadre)
                    tv.Nodes.Add(nodoPadre)
                End If
            Loop
            CN.Close()
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Function
    Sub dameSqlRegistro(ByVal cnreg As SqlConnection, ByVal NombreCampo As String, ByVal NombreTabla As String, ByRef tn As TreeNode)
        Try
            Dim i As Integer = 0

            Dim nodohijo As TreeNode
            Dim query As String
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            cnreg.Open()
            NombreCampo = dameNombreCampoformalSQL(NombreCampo)
            query = "select " & NombreCampo & " from " & NombreTabla
            cmd = New SqlCommand(query, cnreg)
            dr = cmd.ExecuteReader

            Do While dr.Read
                nodohijo = New TreeNode(CStr(dr(0))) 'cojo el primer y unico campo de cada registro
                tn.Nodes.Add(nodohijo)

            Loop
            cnreg.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub dameSqlcampo(ByVal cn As SqlConnection, ByVal nombreTabla As String, ByRef tn As TreeNode)
        Try
            Dim cnnew As New SqlConnection
            Dim i As Integer


            Dim nodohijo As TreeNode
            Dim query As String
            Dim cmd As SqlCommand
            Dim dr As SqlDataReader
            cnnew.ConnectionString = cn.ConnectionString 'por si hay enlazado una conexion anterior
            cn.Open()
            nombreTabla = dameNombreCampoformalSQL(nombreTabla)
            query = "select * from " & nombreTabla
            cmd = New SqlCommand(query, cn)
            dr = cmd.ExecuteReader

            For i = 0 To dr.FieldCount - 1
                nodohijo = New TreeNode(dr.GetName(i))

                dameSqlRegistro(cnnew, dr.GetName(i), nombreTabla, nodohijo)
                tn.Nodes.Add(nodohijo)
            Next
            cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Function dameNombreCampoformalSQL(ByVal nomcampo As String) As String
        Dim empaqueta As Boolean = False
        For Each car As Char In nomcampo
            If Char.IsSeparator(car) Or Char.IsPunctuation(car) Then
                empaqueta = True
            End If
        Next
        If empaqueta Then
            nomcampo = "[" & nomcampo & "]"
        End If
        Return nomcampo
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

End Class