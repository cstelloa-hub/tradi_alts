Attribute VB_Name = "mod_FMS"
Option Explicit

'=====================================================================
' MODULO: mod_FMS  |  para pegar en Fondos Tradicionales.xlsm
'---------------------------------------------------------------------
' Fuente: la carpeta CARPETA_FMS donde estan todos los archivos
'         FMS_AAAAMMDD.xlsx (uno por cierre).
' De cada MES toma el archivo con la fecha mas alta (el cierre del
' mes: si hay FMS diarios, agarra el ultimo dia disponible).
' Extrae SOLO Fondo 1 y Fondo 2 (donde viven los trad):
'   - Posiciones por Codigo SBS (col D), Val_total S/ (col J),
'     sumando si un fondo aparece en varias lineas
'   - AUM total de la cartera (fila "1.1.1 VALOR DE LA CARTERA")
' Y pega VALORES en la hoja "FMS" de este workbook (la crea sola la
' primera vez): Fondo 1 filas 5-16 + AUM fila 17, Fondo 2 filas
' 22-33 + AUM fila 34, una columna por mes.
'
' DOS FORMAS DE CORRERLA:
'   ImportarFMS_UltimoMes -> solo el cierre mas reciente (rutina mensual)
'   ImportarFMS           -> toda la historia (reconstruye todo, idempotente)
'=====================================================================

'--------------------- CONFIGURACION (editar aqui) -------------------
Private Const CARPETA_FMS As String = _
    "\\PFLMVIPR1FS0222\Area de Inversiones\2. Inversiones Alternativas" & _
    "\05. Fondos Tradicionales\FMS\"
Private Const HOJA_FMS As String = ""      ' nombre de la hoja dentro de cada
                                           ' FMS_AAAAMMDD; vacio = primera hoja
Private Const HOJA_DEST As String = "FMS"

'--------------------- LAYOUT HOJA DESTINO ---------------------------
Private Const COL_MES_INI As Long = 5      ' col E = primer mes
Private Const COL_SBS As Long = 3          ' col C = Codigo SBS
Private Const P_FECHA1 As Long = 4         ' fila fechas bloque Fondo 1
Private Const P_F1_INI As Long = 5
Private Const P_F1_FIN As Long = 16
Private Const P_F1_AUM As Long = 17
Private Const P_FECHA2 As Long = 21        ' fila fechas bloque Fondo 2
Private Const P_F2_INI As Long = 22
Private Const P_F2_FIN As Long = 33
Private Const P_F2_AUM As Long = 34

'--------------------- LAYOUT ARCHIVOS FMS ---------------------------
Private Const FMS_FONDO As Long = 1        ' col A: 01=F1, 02=F2 (03/09 se ignoran)
Private Const FMS_SBS As Long = 4          ' col D: Cod_sbs
Private Const FMS_VAL As Long = 10         ' col J: Val_total (S/)
Private Const MARCA_AUM As String = "1.1.1"

'--------------------- CATALOGO DE FONDOS ----------------------------
Private Function Catalogo() As Variant
    Catalogo = Array( _
        Array("HMC - Credito Peru II (PEN)", "4975141HMPSX"), _
        Array("Fondo Credicorp Deuda Soles", "4911011FCDS1"), _
        Array("Fondo Credicorp Deuda Soles II", "4911011FCDS2"), _
        Array("PEPCO II", "4911021CORF2"), _
        Array("Fondo Credicorp Deuda Titulizada", "4911011CCDTI"), _
        Array("Fondo HMC Credito Peru III (PEN)", "4975141HM3SX"), _
        Array("BD Capital - Senior Loans Clase C", "4974252SCBDC"), _
        Array("HMC - Credito Peru II (USD)", "4975142HMPDX"), _
        Array("LV FIAFE", "4911102LINAE"), _
        Array("BD Capital - Senior Loans 2", "4974252S2BDC"), _
        Array("Fondo HMC Credito Peru III (USD)", "4975142HM3DX"), _
        Array("Moneda Patria Peru Fixed Income", "4977602MPAFI"))
End Function

'=====================================================================
' ENTRADAS PUBLICAS
'=====================================================================
Public Sub ImportarFMS()
    Importar False
End Sub

Public Sub ImportarFMS_UltimoMes()
    Importar True
End Sub

'=====================================================================
Private Sub Importar(soloUltimo As Boolean)
    Dim wsD As Worksheet, informe As String
    Dim meses() As String, archivos() As String, n As Long, i As Long

    On Error GoTo Manejo
    Set wsD = ObtenerHojaDestino()

    n = ListarCierres(meses, archivos)
    If n = 0 Then
        MsgBox "No encontre archivos FMS_AAAAMMDD en:" & vbCrLf & CARPETA_FMS & _
               vbCrLf & "Revisa CARPETA_FMS al inicio del modulo mod_FMS.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim iIni As Long
    iIni = IIf(soloUltimo, n, 1)
    For i = iIni To n
        CargarArchivo wsD, meses(i), archivos(i), informe
    Next i
    Application.Calculate
    ActualizarGraficosMeses

Limpieza:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If informe = vbNullString Then informe = "Nada que cargar."
    MsgBox informe, vbInformation, "Importar FMS (Fondo 1 y 2)"
    Exit Sub

Manejo:
    informe = "ERROR " & Err.Number & ": " & Err.Description & vbCrLf & informe
    Resume Limpieza
End Sub

'--- Escanea la carpeta y devuelve, por mes, el archivo de cierre ----
'    (la fecha AAAAMMDD mas alta de cada AAAAMM), orden ascendente
Private Function ListarCierres(ByRef meses() As String, _
                               ByRef archivos() As String) As Long
    Dim dArch As Object, dDia As Object
    Set dArch = CreateObject("Scripting.Dictionary")
    Set dDia = CreateObject("Scripting.Dictionary")

    Dim carpeta As String, f As String, f8 As String
    Dim ym As String, mm As Long, dd As Long
    carpeta = CARPETA_FMS
    If Right$(carpeta, 1) <> "\" Then carpeta = carpeta & "\"

    f = Dir$(carpeta & "FMS_*.xls*")
    Do While f <> vbNullString
        If UCase$(f) Like "FMS_########.XLS*" Then
            f8 = Mid$(f, 5, 8)                    ' AAAAMMDD
            ym = Left$(f8, 6)
            mm = CLng(Mid$(f8, 5, 2))
            dd = CLng(Right$(f8, 2))
            If mm >= 1 And mm <= 12 And dd >= 1 And dd <= 31 Then
                If Not dDia.Exists(ym) Then
                    dArch(ym) = f: dDia(ym) = dd
                ElseIf dd > dDia(ym) Then         ' se queda con el cierre
                    dArch(ym) = f: dDia(ym) = dd
                End If
            End If
        End If
        f = Dir$()
    Loop

    Dim n As Long, i As Long, j As Long, k As Variant, tmp As String
    n = dArch.Count
    If n = 0 Then Exit Function
    ReDim meses(1 To n): ReDim archivos(1 To n)
    i = 0
    For Each k In dArch.Keys
        i = i + 1: meses(i) = CStr(k)
    Next k
    For i = 1 To n - 1                            ' orden ascendente por AAAAMM
        For j = i + 1 To n
            If meses(j) < meses(i) Then
                tmp = meses(i): meses(i) = meses(j): meses(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To n
        archivos(i) = carpeta & dArch(meses(i))
    Next i
    ListarCierres = n
End Function

'--- Abre un FMS_AAAAMMDD, lee F1/F2 y pega la columna del mes -------
Private Sub CargarArchivo(wsD As Worksheet, ym As String, ruta As String, _
                          ByRef informe As String)
    Dim wbF As Workbook, wsF As Worksheet
    Dim nombre As String, abriYo As Boolean

    nombre = Mid$(ruta, InStrRev(ruta, "\") + 1)
    On Error Resume Next
    Set wbF = Workbooks(nombre)                   ' ya abierto?
    On Error GoTo 0
    If wbF Is Nothing Then
        Set wbF = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0)
        abriYo = True
    End If

    If HOJA_FMS <> vbNullString Then
        Set wsF = wbF.Worksheets(HOJA_FMS)
    Else
        Set wsF = wbF.Worksheets(1)
    End If

    Dim dPos As Object, dAUM As Object
    Set dPos = CreateObject("Scripting.Dictionary")
    Set dAUM = CreateObject("Scripting.Dictionary")
    LeerFMS wsF, dPos, dAUM

    If abriYo Then wbF.Close SaveChanges:=False

    Dim fecha As Date, col As Long
    fecha = DateSerial(CLng(Left$(ym, 4)), CLng(Right$(ym, 2)), 1)
    col = ColDeFecha(wsD, fecha)
    If col = 0 Then col = CrearColumna(wsD, fecha)

    Dim faltan As String, k As Long
    k = EscribirBloque(wsD, col, P_F1_INI, P_F1_FIN, P_F1_AUM, "01", dPos, dAUM, faltan)
    k = k + EscribirBloque(wsD, col, P_F2_INI, P_F2_FIN, P_F2_AUM, "02", dPos, dAUM, faltan)

    informe = informe & nombre & " -> " & Format$(fecha, "mmm-yy") & ": " & _
              k & " posiciones" & _
              IIf(dAUM.Exists("01"), ", AUM F1 ok", ", AUM F1 NO ENCONTRADO") & _
              IIf(dAUM.Exists("02"), ", AUM F2 ok", ", AUM F2 NO ENCONTRADO")
    If faltan <> vbNullString Then informe = informe & " | sin posicion (0): " & faltan
    informe = informe & vbCrLf
End Sub

'--- Lee una hoja FMS: posiciones (fondo|SBS -> valor) y AUM ---------
'    Solo se usan luego los fondos 01 y 02; el resto se descarta solo
Private Sub LeerFMS(wsF As Worksheet, dPos As Object, dAUM As Object)
    Dim celda As Range, ultFila As Long, r As Long
    Set celda = wsF.Cells.Find(What:="*", LookIn:=xlFormulas, _
                               SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If celda Is Nothing Then Exit Sub
    ultFila = celda.Row

    Dim txtD As String, fondo As String, clave As String, v As Variant
    For r = 2 To ultFila
        txtD = Trim$(CStr(wsF.Cells(r, FMS_SBS).Value))
        If InStr(1, txtD, MARCA_AUM, vbTextCompare) > 0 Then
            fondo = NormFondo(wsF.Cells(r, FMS_FONDO).Value)
            v = wsF.Cells(r, wsF.Columns.Count).End(xlToLeft).Value
            If fondo <> vbNullString And IsNumeric(v) Then dAUM(fondo) = CDbl(v)
        ElseIf txtD <> vbNullString Then
            v = wsF.Cells(r, FMS_VAL).Value
            If IsNumeric(v) And Not IsEmpty(v) Then
                fondo = NormFondo(wsF.Cells(r, FMS_FONDO).Value)
                If fondo = "01" Or fondo = "02" Then      ' SOLO Fondo 1 y 2
                    clave = fondo & "|" & UCase$(txtD)
                    dPos(clave) = dPos(clave) + CDbl(v)   ' acumula duplicados
                End If
            End If
        End If
    Next r
End Sub

'--- Escribe un bloque; devuelve # posiciones escritas ---------------
Private Function EscribirBloque(wsD As Worksheet, col As Long, fIni As Long, _
                                fFin As Long, fAUM As Long, fondo As String, _
                                dPos As Object, dAUM As Object, _
                                ByRef faltan As String) As Long
    Dim r As Long, sbs As String, clave As String, k As Long
    For r = fIni To fFin
        sbs = UCase$(Trim$(CStr(wsD.Cells(r, COL_SBS).Value)))
        If sbs <> vbNullString Then
            clave = fondo & "|" & sbs
            If dPos.Exists(clave) Then
                wsD.Cells(r, col).Value = dPos(clave)
                k = k + 1
            Else
                wsD.Cells(r, col).Value = 0
                faltan = faltan & IIf(faltan = vbNullString, "", ", ") & _
                         Trim$(CStr(wsD.Cells(r, 2).Value)) & " (F" & Val(fondo) & ")"
            End If
        End If
    Next r
    If dAUM.Exists(fondo) Then wsD.Cells(fAUM, col).Value = dAUM(fondo)
    EscribirBloque = k
End Function

'--- Crea (si hace falta) y devuelve la hoja destino -----------------
Private Function ObtenerHojaDestino() As Worksheet
    Dim ws As Worksheet, cat As Variant, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_DEST)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set ObtenerHojaDestino = ws
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = HOJA_DEST
    cat = Catalogo()

    ws.Cells(1, 1).Value = "Datos FMS (Fondo 1 y 2) - pegado por macro ImportarFMS (NO EDITAR)"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(3, 2).Value = "VALORES FONDO 1 (S/)"
    ws.Cells(20, 2).Value = "VALORES FONDO 2 (S/)"
    ws.Cells(3, 2).Font.Bold = True: ws.Cells(20, 2).Font.Bold = True

    EncabezadoBloque ws, P_FECHA1
    EncabezadoBloque ws, P_FECHA2
    For i = 0 To UBound(cat)
        ws.Cells(P_F1_INI + i, 2).Value = cat(i)(0)
        ws.Cells(P_F1_INI + i, COL_SBS).Value = cat(i)(1)
        ws.Cells(P_F2_INI + i, 2).Value = cat(i)(0)
        ws.Cells(P_F2_INI + i, COL_SBS).Value = cat(i)(1)
    Next i
    ws.Cells(P_F1_AUM, 2).Value = "AUM Total Fondo 1"
    ws.Cells(P_F2_AUM, 2).Value = "AUM Total Fondo 2"
    ws.Cells(P_F1_AUM, 2).Font.Bold = True
    ws.Cells(P_F2_AUM, 2).Font.Bold = True

    ws.Columns(2).ColumnWidth = 32
    ws.Columns(COL_SBS).ColumnWidth = 14
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
    Set ObtenerHojaDestino = ws
End Function

Private Sub EncabezadoBloque(ws As Worksheet, fila As Long)
    ws.Cells(fila, 2).Value = "Fund Name"
    ws.Cells(fila, COL_SBS).Value = "Codigo SBS"
    With ws.Range(ws.Cells(fila, 2), ws.Cells(fila, COL_SBS))
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
End Sub

'------------------------- AUXILIARES --------------------------------
Private Function ColDeFecha(wsD As Worksheet, fecha As Date) As Long
    Dim c As Long, v As Variant
    For c = COL_MES_INI To 300
        v = wsD.Cells(P_FECHA1, c).Value
        If IsEmpty(v) Then Exit For
        If IsDate(v) Then
            If CDate(v) = fecha Then
                ColDeFecha = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function CrearColumna(wsD As Worksheet, fecha As Date) As Long
    Dim c As Long
    c = COL_MES_INI
    Do While Not IsEmpty(wsD.Cells(P_FECHA1, c).Value)
        c = c + 1
    Loop
    wsD.Cells(P_FECHA1, c).Value = fecha
    wsD.Cells(P_FECHA1, c).NumberFormat = "mmm-yy"
    wsD.Cells(P_FECHA1, c).Font.Bold = True
    wsD.Cells(P_FECHA2, c).Value = fecha
    wsD.Cells(P_FECHA2, c).NumberFormat = "mmm-yy"
    wsD.Cells(P_FECHA2, c).Font.Bold = True
    wsD.Columns(c).ColumnWidth = 13
    wsD.Range(wsD.Cells(P_F1_INI, c), wsD.Cells(P_F1_AUM, c)).NumberFormat = "#,##0"
    wsD.Range(wsD.Cells(P_F2_INI, c), wsD.Cells(P_F2_AUM, c)).NumberFormat = "#,##0"
    CrearColumna = c
End Function

Private Function NormFondo(v As Variant) As String
    On Error Resume Next
    If Len(Trim$(CStr(v))) > 0 Then
        If IsNumeric(v) Then NormFondo = Format$(CLng(v), "00")
    End If
End Function

'=====================================================================
' CREARMONITOR: construye la hoja "Retorno por meses" (formulas)
'---------------------------------------------------------------------
'   - Bloque 1 "Retornos por meses": 12 fondos x meses dic-25..dic-27,
'     jalados de "2. Retornos (2)" con INDICE/COINCIDIR por CODIGO SBS
'     (col B de esa hoja) y por fecha (fila de headers).
'     Mes sin dato en la fuente -> celda vacia (no 0).
'   - Bloque 2 "Base 0": dic-25 = 0% y crece compuesto:
'     (1+acum anterior)*(1+retorno del mes)-1; mes vacio arrastra.
'   - Graficos "YTD Soles" (6 PEN) y "YTD Dolares" (6 USD) desde Base 0.
' Idempotente: si la hoja existe, la reconstruye.
' ActualizarGraficosMeses extiende los graficos al ultimo mes con dato
' (se llama sola al final de ImportarFMS / ImportarFMS_UltimoMes).
'=====================================================================

'------------------- ANCLAS EN "2. Retornos (2)" ---------------------
' Verificar contra tu hoja; si algo esta en otra fila/col, se cambia aqui
Private Const R2_HOJA As String = "2. Retornos (2)"
Private Const R2_COL_SBS As String = "B"       ' col de Codigo SBS
Private Const R2_FILA_FECHAS As Long = 7       ' fila de headers de fecha
Private Const R2_FILA_INI As Long = 8          ' primera fila de fondos
Private Const R2_FILA_FIN As Long = 30         ' holgura: cubre PEN y USD
Private Const R2_COL_INI As String = "F"       ' primera col de datos
Private Const R2_COL_FIN As String = "EQ"      ' ultima col de datos

'------------------- LAYOUT "Retorno por meses" ----------------------
Private Const HOJA_RPM As String = "Retorno por meses"
Private Const RPM_MES_INI As Long = 5          ' col E = dic-25
Private Const RPM_MES_FIN As Long = 29         ' col AC = dic-27
Private Const RPM_FECHA1 As Long = 4           ' headers bloque retornos
Private Const RPM_R_INI As Long = 5
Private Const RPM_R_FIN As Long = 16
Private Const RPM_FECHA2 As Long = 20          ' headers bloque Base 0
Private Const RPM_B_INI As Long = 21
Private Const RPM_B_FIN As Long = 32

Public Sub CrearMonitor()
    Dim ws As Worksheet, cat As Variant
    Dim i As Long, c As Long, f As Date

    Application.ScreenUpdating = False
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(HOJA_RPM).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = HOJA_RPM
    cat = Catalogo()

    ws.Cells(1, 1).Value = "Retorno por meses - generado por macro CrearMonitor " & _
                           "(formulas desde '" & R2_HOJA & "', NO EDITAR)"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(3, 2).Value = "RETORNOS POR MESES"
    ws.Cells(19, 2).Value = "BASE 0 (dic-25 = 0%)"
    ws.Cells(3, 2).Font.Bold = True: ws.Cells(19, 2).Font.Bold = True

    EncabezadoBloque ws, RPM_FECHA1
    EncabezadoBloque ws, RPM_FECHA2
    For c = RPM_MES_INI To RPM_MES_FIN         ' fechas dic-25..dic-27
        f = DateSerial(2025, 12 + (c - RPM_MES_INI), 1)
        ws.Cells(RPM_FECHA1, c).Value = f
        ws.Cells(RPM_FECHA2, c).Value = f
        ws.Cells(RPM_FECHA1, c).NumberFormat = "mmm-yy"
        ws.Cells(RPM_FECHA2, c).NumberFormat = "mmm-yy"
        ws.Cells(RPM_FECHA1, c).Font.Bold = True
        ws.Cells(RPM_FECHA2, c).Font.Bold = True
        ws.Columns(c).ColumnWidth = 9
    Next c

    Dim src As String, idx As String, refRet As String, refPrev As String
    src = "'" & R2_HOJA & "'!"
    For i = 0 To UBound(cat)
        ws.Cells(RPM_R_INI + i, 2).Value = cat(i)(0)
        ws.Cells(RPM_R_INI + i, 3).Value = cat(i)(1)
        ws.Cells(RPM_B_INI + i, 2).Value = cat(i)(0)
        ws.Cells(RPM_B_INI + i, 3).Value = cat(i)(1)

        For c = RPM_MES_INI To RPM_MES_FIN
            ' ---- bloque retornos: INDICE/COINCIDIR por SBS y fecha ----
            idx = "INDEX(" & src & "$" & R2_COL_INI & "$" & R2_FILA_INI & _
                  ":$" & R2_COL_FIN & "$" & R2_FILA_FIN & _
                  ",MATCH($C" & (RPM_R_INI + i) & "," & src & "$" & R2_COL_SBS & _
                  "$" & R2_FILA_INI & ":$" & R2_COL_SBS & "$" & R2_FILA_FIN & ",0)" & _
                  ",MATCH(" & ws.Cells(RPM_FECHA1, c).Address(True, False) & "," & _
                  src & "$" & R2_COL_INI & "$" & R2_FILA_FECHAS & ":$" & _
                  R2_COL_FIN & "$" & R2_FILA_FECHAS & ",0))"
            ws.Cells(RPM_R_INI + i, c).Formula = _
                "=IFERROR(IF(" & idx & "="""",""""," & idx & "),"""")"
            ws.Cells(RPM_R_INI + i, c).NumberFormat = "0.00%"

            ' ---- bloque Base 0 -----------------------------------------
            If c = RPM_MES_INI Then
                ws.Cells(RPM_B_INI + i, c).Value = 0
            Else
                refRet = ws.Cells(RPM_R_INI + i, c).Address(False, False)
                refPrev = ws.Cells(RPM_B_INI + i, c - 1).Address(False, False)
                ws.Cells(RPM_B_INI + i, c).Formula = _
                    "=IF(" & refRet & "=""""," & refPrev & _
                    ",(1+" & refPrev & ")*(1+" & refRet & ")-1)"
            End If
            ws.Cells(RPM_B_INI + i, c).NumberFormat = "0.00%"
        Next c
    Next i

    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 32
    ws.Columns(3).ColumnWidth = 14
    ws.Columns(4).ColumnWidth = 5
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
    ws.Cells(1, 1).Font.Size = 8

    CrearGraficoRPM ws, "YTD Soles", RPM_B_INI, RPM_B_INI + 5, ws.Range("B35")
    CrearGraficoRPM ws, "YTD Dolares", RPM_B_INI + 6, RPM_B_FIN, ws.Range("J35")
    ActualizarGraficosMeses

    Application.ScreenUpdating = True
    MsgBox "Hoja '" & HOJA_RPM & "' creada: retornos jalados de '" & R2_HOJA & _
           "' + Base 0 + graficos.", vbInformation, "Crear Monitor"
End Sub

Private Sub CrearGraficoRPM(ws As Worksheet, titulo As String, _
                            rIni As Long, rFin As Long, celda As Range)
    Dim co As ChartObject, s As Series, r As Long
    Set co = ws.ChartObjects.Add(celda.Left, celda.Top, 420, 240)
    co.Chart.ChartType = xlLine
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = titulo
    For r = rIni To rFin
        Set s = co.Chart.SeriesCollection.NewSeries
        s.Name = "='" & ws.Name & "'!" & ws.Cells(r, 2).Address
        s.XValues = ws.Range(ws.Cells(RPM_FECHA2, RPM_MES_INI), _
                             ws.Cells(RPM_FECHA2, RPM_MES_INI + 1))
        s.Values = ws.Range(ws.Cells(r, RPM_MES_INI), _
                            ws.Cells(r, RPM_MES_INI + 1))
    Next r
End Sub

'--- Extiende los graficos al ultimo mes con retorno -----------------
Public Sub ActualizarGraficosMeses()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_RPM)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastCol As Long, c As Long
    lastCol = RPM_MES_INI + 1
    For c = RPM_MES_FIN To RPM_MES_INI + 1 Step -1
        If Application.WorksheetFunction.Count( _
           ws.Range(ws.Cells(RPM_R_INI, c), ws.Cells(RPM_R_FIN, c))) > 0 Then
            lastCol = c
            Exit For
        End If
    Next c

    Dim co As ChartObject, s As Long, base As Long
    For Each co In ws.ChartObjects
        If InStr(co.Chart.SeriesCollection(1).Formula, _
                 "$B$" & RPM_B_INI & """") > 0 Or _
           InStr(co.Chart.SeriesCollection(1).Formula, _
                 "$B$" & RPM_B_INI & "'") > 0 Or _
           InStr(co.Chart.SeriesCollection(1).Formula, _
                 "$B$" & RPM_B_INI & ",") > 0 Then
            base = RPM_B_INI
        Else
            base = RPM_B_INI + 6
        End If
        For s = 1 To co.Chart.SeriesCollection.Count
            With co.Chart.SeriesCollection(s)
                .XValues = ws.Range(ws.Cells(RPM_FECHA2, RPM_MES_INI), _
                                    ws.Cells(RPM_FECHA2, lastCol))
                .Values = ws.Range(ws.Cells(base + s - 1, RPM_MES_INI), _
                                   ws.Cells(base + s - 1, lastCol))
            End With
        Next s
    Next co
End Sub
