Public Class CalendarioFMC
    'creado por Fernando Mendiz
    '4/07/15 
    '12:09
    'v1.00
    'cambios : inicial
    'Para usar 
    ' en el load asignar FSalida y FVuelta son calendarios
    '         FechaSalida.Text = FSalida.SelectedDate.ToShortDateString
    '         FechaVuelta.Text = FVuelta.SelectedDate.ToShortDateString
    'incluir rangoValidator compareValidator validationSummary
    'Sub New(ByVal dni As Integer)
    '    _dni = dni
    '    Calculaletra()
    'End Sub

    Public errorFMC As String = ""
    'Public RV As New Object
    'Public txtB1 As New Object

    Sub New()

    End Sub
    'Sub New(ByRef rvalidator As Object, ByVal textoB As Object)
    '    RV = rvalidator
    '    txtB1 = textoB
    'End Sub


   

    Public Function ValidaRango(ByRef RV As Object, fechaMIN As DateTime, fechaMAX As DateTime, errorMensajeFMC As String) As String
        'pre: RV debe tener el parametro ControlToValidate enlazado y DateTime.Now.AddDays(1).ToShortDateString 'para mañana
        '   DateTime.Now.AddMonths(1).ToShortDateString() 'para el ultimo mes
        '   en errorMensajeFMC el mensaje si falla el validador
        'post: si RV es un objeto RangeValidator y txtB1 comprueba que su valor sea como minimo mañana y dentro de un mes como maximo
        ' Dim errorFMC As String = ""
        'Dim r As New RangeValidator
        If (compruebaEsObjetoRangeValidator(RV) AndAlso compruebaValidatorEnlazado(RV)) Then
            RV.ErrorMessage = errorMensajeFMC
            RV.MinimumValue = fechaMIN.ToShortDateString
            RV.MaximumValue = fechaMAX.ToShortDateString
            RV.Type = ValidationDataType.Date
            RV.Text = "seleccione una fecha valida"
        Else
            If compruebaEsObjetoRangeValidator(RV) Then
                Me.errorFMC = "no es un rangeValidate lo que has pasado"
            Else
                Me.errorFMC = "range validate no enlazado"
            End If
            '   
        End If

        Return Me.errorFMC
    End Function

    Public Function CreaBotonAtras(ByRef bo As Object) As String
        'pre: este codigo se pone en el page.load
        'post: si es un boton bo lo convierte en un boton que recupera la pagina anterior

        If compruebaEsObjetoButton(bo) Then
            ' Dim b As New Button
            bo.Text = "Atras"
            bo.UseSubmitBehavior = "False"
            bo.OnClientClick = "window.history.back(1); return false"
        Else
            Me.errorFMC = "no es un boton lo que has pasado"
        End If
        Return Me.errorFMC
    End Function

    Public Function CreaBotonAtras(ByRef bo As Object, text As String) As String
        'pre: 
        'post: si es un boton bo lo convierte en un boton que recupera la pagina anterior

        If compruebaEsObjetoButton(bo) Then
            ' Dim b As New Button
            bo.Text = text
            bo.UseSubmitBehavior = "False"
            bo.OnClientClick = "window.history.back(1); return false"

        Else

            Me.errorFMC = "no es un boton lo que has pasado"

        End If
        Return Me.errorFMC
    End Function

    Public Function ValidaRangoMañanaHasta1Mes(ByRef RV As Object) As String
        'pre:
        'post: si RV es un objeto RangeValidator y txtB1 comprueba que su valor sea como minimo mañana y dentro de un mes como maximo
        'Dim errorFMC As String = ""
        If (compruebaEsObjetoRangeValidator(RV) AndAlso compruebaValidatorEnlazado(RV)) Then
            'Dim r As New RangeValidator
            RV.ErrorMessage = "La fecha de salida debe de ser posterior al dia de hoy y dentro del plazo de 30 dias"
            'Me.errorFMC = "La fecha de salida debe de ser posterior al dia de hoy y dentro del plazo de 30 dias"
            RV.MinimumValue = DateTime.Now.AddDays(1).ToShortDateString 'mañana
            RV.MaximumValue = DateTime.Now.AddMonths(1).ToShortDateString 'el ultimo mes
            RV.Type = ValidationDataType.Date
            RV.Text = "seleccione una fecha valida"
        Else
            If Not compruebaEsObjetoRangeValidator(RV) Then
                Me.errorFMC = "no es un rangeValidate lo que has pasado"
            Else
                Me.errorFMC = "range validate no enlazado"
            End If
        End If

        Return Me.errorFMC
    End Function

    Public Function comparaFechaVuelta(ByRef CV As Object, errorMensajeFMC As String, textobox As Object, operador As ValidationCompareOperator) As String
        If (compruebaEsObjetoCompareValidator(CV) AndAlso compruebaValidatorEnlazado(CV)) Then
            'pre: 
            'post:
            'Dim c As New CompareValidator
            'Dim t As New TextBox

            CV.ErrorMessage = errorMensajeFMC
            CV.ControlToCompare = textobox.ID
            CV.Operator = operador
            CV.Type = ValidationDataType.Date
            CV.Text = "seleccione una fecha valida"
        Else
            If Not compruebaEsObjetoCompareValidator(CV) Then
                Me.errorFMC = "no es un CompareValidate lo que has pasado a comparaFechaVuelta"
            Else
                Me.errorFMC = "comparevalidate no enlazado"
            End If

        End If
        Return Me.errorFMC
    End Function


    Public Function comparaFechaVueltaMayorFechaSalida(ByRef cv As Object, textobox As Object) As String
        'pre:  
        'post:  si CV es un objeto compareValidator enlazado con textobox comprueba que su valor mayor que la fecha de texto box
        'Public Function comparaFechaVuelta(ByRef CV As Object, errorMensajeFMC As String, textobox As Object, operador As ValidationCompareOperator) As String
        If (compruebaEsObjetoCompareValidator(cv) AndAlso compruebaValidatorEnlazado(cv)) Then
            'Dim c As New CompareValidator
            cv.ErrorMessage = "la Fecha de vuelta debe ser posterior a la salida"
            cv.ControlToCompare = textobox.ID
            cv.Type = ValidationDataType.Date
            cv.Operator = ValidationCompareOperator.GreaterThan
            cv.Text = "seleccione una fecha valida"
        Else
            If Not compruebaEsObjetoCompareValidator(cv) Then
                Me.errorFMC = "no es un CompareValidate lo que has pasado a comparaFechaVueltaMayorFechaSalida"
            Else
                Me.errorFMC = "comparevalidate no enlazado"
            End If

        End If
        Return Me.errorFMC
    End Function

    Public Function FormateaValidationSumary(vs As Object) As String
        If compruebaEsObjetoCalendario(vs) Then
            'Dim v As New ValidationSummary

            vs.DisplayMode = ValidationSummaryDisplayMode.BulletList
        Else
            Me.errorFMC = "no has pasado un validationSumary"
        End If
        Return Me.errorFMC
    End Function

    Public Function formateaCalendario(ByRef cal As Object, Titulo As String) As String
        'pre: cal es un objeto 
        'post: si cal es un objeto calendario lo formatea
        ' Dim errorFMC As String = ""
        'Dim C As New Calendar
        If compruebaEsObjetoCalendario(cal) Then

            cal.Caption = Titulo
            cal.BackColor = Drawing.Color.White
            cal.BorderColor = Drawing.ColorTranslator.FromHtml("#999999")
            cal.CellPadding = "4"
            cal.DayNameFormat = DayNameFormat.Shortest
            cal.Font.Name = "Verdana"
            cal.Font.Size = 8
            cal.ForeColor = Drawing.Color.Black
            cal.Height = 180
            cal.Width = 200
            cal.DayHeaderStyle.BackColor = Drawing.ColorTranslator.FromHtml("#CCCCCC")
            cal.DayHeaderStyle.BorderColor = Drawing.Color.Black
            cal.DayHeaderStyle.Font.Bold = "True"
            cal.DayHeaderStyle.Font.Size = 7
            cal.NextPrevStyle.VerticalAlign = VerticalAlign.Bottom
            cal.OtherMonthDayStyle.BackColor = Drawing.Color.Gray
            cal.SelectedDayStyle.BackColor = Drawing.ColorTranslator.FromHtml("#666666")
            cal.SelectedDayStyle.Font.Bold = True
            cal.SelectorStyle.BackColor = Drawing.ColorTranslator.FromHtml("#999999")
            cal.TodayDayStyle.BackColor = Drawing.ColorTranslator.FromHtml("#CCCCCC")
            cal.TodayDayStyle.ForeColor = Drawing.Color.Black

        Else
            Me.errorFMC = "no es un calendario lo que has pasado"
        End If
        Return Me.errorFMC
    End Function
    Private Function compruebaEsObjetoCalendario(cal As Object) As Boolean
        Dim tip As Type
        Dim tipoCal = "System.Web.UI.WebControls.Calendar"
        tip = cal.GetType()
        Return (String.Compare(tip.ToString, tipoCal) = 0)
    End Function

    Private Function compruebaEsObjetoRangeValidator(RV As Object) As Boolean
        Dim tip As Type
        Dim tipo = "System.Web.UI.WebControls.RangeValidator"
        'Dim tipo As New RangeValidator
        tip = RV.GetType()
        Return (String.Compare(tip.ToString, tipo) = 0)
    End Function
    Private Function compruebaEsObjetoCompareValidator(CV As Object) As Boolean
        Dim tip As Type
        Dim tipo = "System.Web.UI.WebControls.CompareValidator"
        'Dim tipo As New CompareValidator
        tip = CV.GetType()
        Return (String.Compare(tip.ToString, tipo) = 0)
    End Function

    Private Function compruebaEsObjetoTextBox(txtb As Object) As Boolean
        Dim tip As Type
        Dim tipo = "System.Web.UI.WebControls.TextBox"
        'Dim tipo As New TextBox
        tip = txtb.GetType()
        Return (String.Compare(tip.ToString, tipo) = 0)
    End Function

    Private Function compruebaEsObjetoValidationSummary(txtb As Object) As Boolean
        Dim tip As Type
        Dim tipo = "System.Web.UI.WebControls.ValidationSummary"
        'Dim tipo As New ValidationSummary
        tip = txtb.GetType()
        Return (String.Compare(tip.ToString, tipo) = 0)
    End Function

    Private Function compruebaEsObjetoButton(txtb As Object) As Boolean
        Dim tip As Type
        Dim tipo = "System.Web.UI.WebControls.Button"
        'Dim tipo As New ValidationSummary
        tip = txtb.GetType()
        Return (String.Compare(tip.ToString, tipo) = 0)
    End Function

    Private Function compruebaValidatorEnlazado(v As BaseCompareValidator)
        Return Not String.Compare(v.ControlToValidate, "") = 0 'control no enlazado
    End Function


    'dependencias de la validacion no intrusiva
    'jquery  y aspNet.ScripManager.jquery

    'validacion no intrusiva
    'modificar el web.config
    'añadiendo:
    '<appSetting>
    '<add key="ValidationSettings:UnobstrusiveValidationMode"
    '   value=none />
    '</appSetting>

    ' agregando un modulo Global.asax
    'y en el metodo Application_Start
    'Ponemos:
    ' ValidationSettings.UnobtrusiveValidationMode = System.Web.UI.UnobtrusiveValidationMode.None



    '----------------------
    'si queries mostrar error en tiempo ejecucion crea una variable publica ErrorFmc en el *.aspx.vb
    'descarga el en ella el resultado de cualquier funcion de esta clase
    'pegando abajo esto:
    'If Not String.Compare(ErrorFMC, "") Then
    '      Response.Write(ErrorFMC)
    '  End If
End Class

