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

Private Const HOJA_CALC As String = "Calculos"
Private Const HOJA_RES As String = "Resumen"
Private Const RPMQ As String = "'Retorno por meses'!"


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

    CrearCalculos
    CrearResumen
    Application.Calculate
    ActualizarGraficosMeses

    Application.ScreenUpdating = True
    MsgBox "Monitor creado: hojas '" & HOJA_RPM & "', '" & HOJA_CALC & "' y '" & _
           HOJA_RES & "' construidas con formulas." & vbCrLf & _
           "Falta solo correr ImportarFMS para llenar la hoja FMS.", _
           vbInformation, "Crear Monitor"
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

'=====================================================================
' HOJAS "Calculos" y "Resumen" (las construye CrearMonitor)
'---------------------------------------------------------------------
' Calculos (solo formulas, grid fijo dic-25..dic-27 en cols E..AC):
'   %F1 filas 5-16, %F2 21-32, peso combinado 37-48,
'   contribucion (peso combinado mes anterior x retorno) 53-64,
'   MTD ponderado 65-67, normalizado MTD 71-73 y YTD 74-76.
'   Toda referencia a FMS es por FECHA (INDEX/MATCH), asi que un mes
'   sin archivo en la carpeta simplemente queda vacio, sin descuadrar.
' Resumen: selectores de fecha y fondo, KPIs MTD/YTD por moneda,
'   tablas PEN y USD (anos desde "2. Retornos (2)", YTD/1M/3M/6M,
'   posicion MM y % de F1/F2) con heatmap de formato condicional.
'=====================================================================
Private Function IdxFMS(fila As Long, filaFecha As Long, refFecha As String) As String
    IdxFMS = "INDEX(FMS!$E$" & fila & ":$BZ$" & fila & ",MATCH(" & refFecha & _
             ",FMS!$E$" & filaFecha & ":$BZ$" & filaFecha & ",0))"
End Function

Private Function IdxFMSBlq(f1 As Long, f2 As Long, filaFecha As Long, _
                           refFecha As String) As String
    IdxFMSBlq = "INDEX(FMS!$E$" & f1 & ":$BZ$" & f2 & ",0,MATCH(" & refFecha & _
                ",FMS!$E$" & filaFecha & ":$BZ$" & filaFecha & ",0))"
End Function

Public Sub CrearCalculos()
    Dim ws As Worksheet, cat As Variant
    Dim i As Long, c As Long, L As String, Lp As String, fRef As String, fPrev As String
    Dim pos As String, aum As String, pes As String, ret As String

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(HOJA_CALC).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = HOJA_CALC
    cat = Catalogo()

    ws.Cells(1, 1).Value = "Calculos - generado por macro CrearMonitor (solo formulas, NO EDITAR)"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(3, 2).Value = "PESO % FONDO 1": ws.Cells(19, 2).Value = "PESO % FONDO 2"
    ws.Cells(35, 2).Value = "PESO % COMBINADO (F1+F2)"
    ws.Cells(51, 2).Value = "CONTRIBUCION MENSUAL (peso combinado mes anterior x retorno)"
    ws.Cells(69, 2).Value = "NORMALIZADO (fondos listados = 100%)"
    Dim rl As Variant
    For Each rl In Array(3, 19, 35, 51, 69)
        ws.Cells(CLng(rl), 2).Font.Bold = True
    Next rl

    For c = RPM_MES_INI To RPM_MES_FIN         ' fila 4: grid fijo de fechas
        ws.Cells(4, c).Value = DateSerial(2025, 12 + (c - RPM_MES_INI), 1)
        ws.Cells(4, c).NumberFormat = "mmm-yy"
        ws.Cells(4, c).Font.Bold = True
        ws.Columns(c).ColumnWidth = 9
    Next c

    For i = 0 To UBound(cat)
        Dim fN As Variant
        For Each fN In Array(5, 21, 37, 53)
            ws.Cells(CLng(fN) + i, 2).Value = cat(i)(0)
        Next fN
        For c = RPM_MES_INI To RPM_MES_FIN
            L = Split(ws.Cells(1, c).Address, "$")(1)
            fRef = L & "$4"
            ' --- %F1 / %F2 / combinado -----------------------------------
            pos = IdxFMS(5 + i, 4, fRef): aum = IdxFMS(17, 4, fRef)
            ws.Cells(5 + i, c).Formula = "=IFERROR(" & pos & "/" & aum & ","""")"
            pos = IdxFMS(22 + i, 21, fRef): aum = IdxFMS(34, 21, fRef)
            ws.Cells(21 + i, c).Formula = "=IFERROR(" & pos & "/" & aum & ","""")"
            ws.Cells(37 + i, c).Formula = "=IFERROR((" & IdxFMS(5 + i, 4, fRef) & _
                "+" & IdxFMS(22 + i, 21, fRef) & ")/(" & IdxFMS(17, 4, fRef) & _
                "+" & IdxFMS(34, 21, fRef) & "),"""")"
            ws.Cells(5 + i, c).NumberFormat = "0.00%"
            ws.Cells(21 + i, c).NumberFormat = "0.00%"
            ws.Cells(37 + i, c).NumberFormat = "0.00%"
            ' --- contribucion (desde ene-26) -----------------------------
            If c > RPM_MES_INI Then
                Lp = Split(ws.Cells(1, c - 1).Address, "$")(1)
                ws.Cells(53 + i, c).Formula = "=IF(OR(" & RPMQ & L & (RPM_R_INI + i) & _
                    "=""""," & Lp & (37 + i) & "=""""),""""," & Lp & (37 + i) & _
                    "*" & RPMQ & L & (RPM_R_INI + i) & ")"
                ws.Cells(53 + i, c).NumberFormat = "0.000%"
            End If
        Next c
    Next i

    ' --- MTD ponderado + normalizado -----------------------------------
    Dim lbl As Variant, k As Long
    lbl = Array(Array("MTD Ponderado PEN", 53, 58), Array("MTD Ponderado USD", 59, 64), _
                Array("MTD Ponderado Total", 53, 64))
    For k = 0 To 2
        ws.Cells(65 + k, 2).Value = lbl(k)(0): ws.Cells(65 + k, 2).Font.Bold = True
    Next k
    Dim nrm As Variant
    nrm = Array(Array("Normalizado MTD PEN", 5, 10, 22, 27, 5, 10), _
                Array("Normalizado MTD USD", 11, 16, 28, 33, 11, 16), _
                Array("Normalizado MTD Total", 5, 16, 22, 33, 5, 16))
    For k = 0 To 2
        ws.Cells(71 + k, 2).Value = nrm(k)(0): ws.Cells(71 + k, 2).Font.Bold = True
        ws.Cells(74 + k, 2).Value = Replace(CStr(nrm(k)(0)), "MTD", "YTD")
        ws.Cells(74 + k, 2).Font.Bold = True
    Next k

    For c = RPM_MES_INI + 1 To RPM_MES_FIN
        L = Split(ws.Cells(1, c).Address, "$")(1)
        Lp = Split(ws.Cells(1, c - 1).Address, "$")(1)
        fPrev = Lp & "$4"
        For k = 0 To 2
            ws.Cells(65 + k, c).Formula = "=IF(COUNT(" & L & lbl(k)(1) & ":" & _
                L & lbl(k)(2) & ")=0,"""",SUM(" & L & lbl(k)(1) & ":" & L & lbl(k)(2) & "))"
            ws.Cells(65 + k, c).NumberFormat = "0.00%"
            pes = "(" & IdxFMSBlq(CLng(nrm(k)(1)), CLng(nrm(k)(2)), 4, fPrev) & _
                  "+" & IdxFMSBlq(CLng(nrm(k)(3)), CLng(nrm(k)(4)), 21, fPrev) & ")"
            ret = RPMQ & L & nrm(k)(5) & ":" & L & nrm(k)(6)
            ws.Cells(71 + k, c).Formula = "=IFERROR(SUMPRODUCT(" & pes & ",N(" & ret & _
                "))/SUMPRODUCT(" & pes & "*(" & ret & "<>"""")),"""")"
            ws.Cells(71 + k, c).NumberFormat = "0.00%"
            ws.Cells(74 + k, c).Formula = "=IF(" & L & (71 + k) & "="""",""""," & _
                "IF(OR(" & Lp & (74 + k) & "="""",YEAR(" & L & "$4)<>YEAR(" & Lp & "$4))," & _
                L & (71 + k) & ",(1+" & Lp & (74 + k) & ")*(1+" & L & (71 + k) & ")-1))"
            ws.Cells(74 + k, c).NumberFormat = "0.00%"
        Next k
    Next c

    ws.Columns(1).ColumnWidth = 2: ws.Columns(2).ColumnWidth = 34
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
End Sub

Public Sub CrearResumen()
    Dim ws As Worksheet, cat As Variant, i As Long, k As Long
    Dim src As String, rr As Long, r As Long, fi As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(HOJA_RES).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = HOJA_RES
    cat = Catalogo()
    src = "'" & R2_HOJA & "'!"

    ws.Cells(1, 1).Value = "Fondos Tradicionales - Resumen"
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 10

    ' --- selectores -----------------------------------------------------
    ws.Range("B3").Value = "Fecha de reporte:": ws.Range("B3").Font.Bold = True
    ws.Range("F3").Value = "Fondo:": ws.Range("F3").Font.Bold = True
    With ws.Range("C3").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="='" & HOJA_RPM & "'!$F$4:$AC$4"
    End With
    With ws.Range("G3").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="='" & HOJA_RPM & "'!$B$5:$B$16"
    End With
    ws.Range("C3").NumberFormat = "mmm-yy"
    ws.Range("C3").Interior.Color = RGB(255, 242, 204)
    ws.Range("G3").Interior.Color = RGB(255, 242, 204)
    ws.Range("G3").Value = cat(0)(0)
    ws.Range("P3").Formula = "=IFERROR(MATCH($C$3,'" & HOJA_RPM & _
                             "'!$E$4:$AC$4,0),"""")"
    ws.Range("P3").Font.Color = RGB(217, 217, 217)

    ' C3 inicial = ultimo mes con retorno
    Dim wsR As Worksheet, lastCol As Long, c As Long
    Set wsR = ThisWorkbook.Worksheets(HOJA_RPM)
    lastCol = RPM_MES_INI + 1
    For c = RPM_MES_FIN To RPM_MES_INI + 1 Step -1
        If Application.WorksheetFunction.Count( _
           wsR.Range(wsR.Cells(RPM_R_INI, c), wsR.Cells(RPM_R_FIN, c))) > 0 Then
            lastCol = c: Exit For
        End If
    Next c
    ws.Range("C3").Value = wsR.Cells(RPM_FECHA1, lastCol).Value

    ' --- KPIs -----------------------------------------------------------
    ws.Range("B5").Value = "Moneda": ws.Range("C5").Value = "MTD": ws.Range("D5").Value = "YTD"
    EncabezadoFila ws, 5, 2, 4
    Dim mon As Variant: mon = Array("PEN", "USD", "Total")
    For k = 0 To 2
        ws.Cells(6 + k, 2).Value = mon(k): ws.Cells(6 + k, 2).Font.Bold = True
        ws.Cells(6 + k, 3).Formula = "=IF($P$3="""","""",IFERROR(INDEX(" & _
            HOJA_CALC & "!$E$" & (71 + k) & ":$AC$" & (71 + k) & ",$P$3),""""))"
        ws.Cells(6 + k, 4).Formula = "=IF($P$3="""","""",IFERROR(INDEX(" & _
            HOJA_CALC & "!$E$" & (74 + k) & ":$AC$" & (74 + k) & ",$P$3),""""))"
        ws.Cells(6 + k, 3).NumberFormat = "0.00%": ws.Cells(6 + k, 4).NumberFormat = "0.00%"
    Next k
    ws.Range("F5").Value = "Fondo seleccionado": ws.Range("G5").Value = "MTD"
    ws.Range("H5").Value = "Acum. dic-25"
    EncabezadoFila ws, 5, 6, 8
    ws.Range("F6").Formula = "=$G$3"
    ws.Range("G6").Formula = "=IF($P$3="""","""",IFERROR(INDEX('" & HOJA_RPM & _
        "'!$E$5:$AC$16,MATCH($G$3,'" & HOJA_RPM & "'!$B$5:$B$16,0),$P$3),""""))"
    ws.Range("H6").Formula = "=IF($P$3="""","""",IFERROR(INDEX('" & HOJA_RPM & _
        "'!$E$21:$AC$32,MATCH($G$3,'" & HOJA_RPM & "'!$B$5:$B$16,0),$P$3),""""))"
    ws.Range("G6:H6").NumberFormat = "0.00%"

    ' --- tablas PEN y USD ----------------------------------------------
    Dim blq As Variant
    blq = Array(Array("FONDOS PEN", 0, 13), Array("FONDOS USD", 6, 23))
    Dim b As Long
    For b = 0 To 1
        Dim idx0 As Long, rIni As Long
        idx0 = blq(b)(1): rIni = blq(b)(2)
        ws.Cells(rIni - 2, 2).Value = blq(b)(0): ws.Cells(rIni - 2, 2).Font.Bold = True
        Dim hdrs As Variant
        hdrs = Array("Instrumento", "2023", "2024", "2025", "YTD", "1M", "3M", "6M", _
                     "Pos F1 (MM)", "Pos F2 (MM)", "% F1", "% F2")
        For k = 0 To UBound(hdrs)
            ws.Cells(rIni - 1, 2 + k).Value = hdrs(k)
        Next k
        EncabezadoFila ws, rIni - 1, 2, 13

        For i = 0 To 5
            fi = idx0 + i: r = rIni + i: rr = RPM_R_INI + fi
            ws.Cells(r, 2).Formula = "='" & HOJA_RPM & "'!B" & rr
            For k = 0 To 2                     ' 2023 / 2024 / 2025
                Dim yr As String: yr = CStr(2023 + k)
                ws.Cells(r, 3 + k).Formula = "=IFERROR(INDEX(" & src & "$" & _
                    R2_COL_INI & "$" & R2_FILA_INI & ":$" & R2_COL_FIN & "$" & R2_FILA_FIN & _
                    ",MATCH('" & HOJA_RPM & "'!$C$" & rr & "," & src & "$" & R2_COL_SBS & "$" & _
                    R2_FILA_INI & ":$" & R2_COL_SBS & "$" & R2_FILA_FIN & ",0),MATCH(" & yr & _
                    "," & src & "$" & R2_COL_INI & "$" & R2_FILA_FECHAS & ":$" & R2_COL_FIN & _
                    "$" & R2_FILA_FECHAS & ",0)),IFERROR(INDEX(" & src & "$" & R2_COL_INI & _
                    "$" & R2_FILA_INI & ":$" & R2_COL_FIN & "$" & R2_FILA_FIN & ",MATCH('" & _
                    HOJA_RPM & "'!$C$" & rr & "," & src & "$" & R2_COL_SBS & "$" & R2_FILA_INI & _
                    ":$" & R2_COL_SBS & "$" & R2_FILA_FIN & ",0),MATCH(""" & yr & """," & src & _
                    "$" & R2_COL_INI & "$" & R2_FILA_FECHAS & ":$" & R2_COL_FIN & "$" & _
                    R2_FILA_FECHAS & ",0)),""""))"
            Next k
            ws.Cells(r, 6).Formula = "=IF($P$3="""","""",IFERROR(EXP(SUMPRODUCT(LN(1+" & _
                "N(OFFSET('" & HOJA_RPM & "'!$D$4," & (fi + 1) & _
                ",$P$3-MONTH($C$3)+1,1,MONTH($C$3))))))-1,""""))"
            ws.Cells(r, 7).Formula = "=IF($P$3="""","""",IF(INDEX('" & HOJA_RPM & _
                "'!$E$" & rr & ":$AC$" & rr & ",$P$3)="""","""",INDEX('" & HOJA_RPM & _
                "'!$E$" & rr & ":$AC$" & rr & ",$P$3)))"
            ws.Cells(r, 8).Formula = "=IF(OR($P$3="""",$P$3<3),"""",IFERROR(EXP(" & _
                "SUMPRODUCT(LN(1+N(OFFSET('" & HOJA_RPM & "'!$D$4," & (fi + 1) & _
                ",$P$3-2,1,3)))))-1,""""))"
            ws.Cells(r, 9).Formula = "=IF(OR($P$3="""",$P$3<6),"""",IFERROR(EXP(" & _
                "SUMPRODUCT(LN(1+N(OFFSET('" & HOJA_RPM & "'!$D$4," & (fi + 1) & _
                ",$P$3-5,1,6)))))-1,""""))"
            ws.Cells(r, 10).Formula = "=IFERROR(" & IdxFMS(5 + fi, 4, "$C$3") & _
                "/1000000,"""")"
            ws.Cells(r, 11).Formula = "=IFERROR(" & IdxFMS(22 + fi, 21, "$C$3") & _
                "/1000000,"""")"
            ws.Cells(r, 12).Formula = "=IFERROR(" & IdxFMS(5 + fi, 4, "$C$3") & "/" & _
                IdxFMS(17, 4, "$C$3") & ","""")"
            ws.Cells(r, 13).Formula = "=IFERROR(" & IdxFMS(22 + fi, 21, "$C$3") & "/" & _
                IdxFMS(34, 21, "$C$3") & ","""")"
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 9)).NumberFormat = "0.00%"
            ws.Range(ws.Cells(r, 10), ws.Cells(r, 11)).NumberFormat = "#,##0.0"
            ws.Range(ws.Cells(r, 12), ws.Cells(r, 13)).NumberFormat = "0.00%"
        Next i
        ' heatmap 3 colores sobre retornos
        With ws.Range(ws.Cells(rIni, 3), ws.Cells(rIni + 5, 9)).FormatConditions.AddColorScale(3)
            .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
            .ColorScaleCriteria(2).Type = xlConditionValuePercentile
            .ColorScaleCriteria(2).Value = 50
            .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
            .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
        End With
    Next b

    ws.Columns(1).ColumnWidth = 2: ws.Columns(2).ColumnWidth = 32
    For k = 3 To 13
        ws.Columns(k).ColumnWidth = 10
    Next k
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
    ws.Cells(1, 1).Font.Size = 10
End Sub

Private Sub EncabezadoFila(ws As Worksheet, fila As Long, c1 As Long, c2 As Long)
    With ws.Range(ws.Cells(fila, c1), ws.Cells(fila, c2))
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
End Sub
