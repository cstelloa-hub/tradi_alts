Attribute VB_Name = "mod_MonitorDia"
Option Explicit

'=====================================================================
' MODULO: mod_MonitorDia  |  nuevo libro "Monitor Tradicionales (Dia).xlsm"
'---------------------------------------------------------------------
' Todo por DIA DE MARCA. El motor es Marcas.xlsx (Var Adj): retorno del
' fondo, retorno del bloque y contribucion a F1/F2 salen de la marca y
' de la fecha en que entro a cartera (Dia). No hay retornos por EEFF.
'
'  Capa 1 - BASE BRUTA (vertical, solo la macro escribe, sin formulas)
'    BD_FMS     Fecha x Fondo AFP (F1/F2) x instrumento: Valor S/, AUM.
'               Carga incremental de los FMS_AAAAMMDD.
'    BD_Marcas  una fila por marca tal como viene de Marcas.xlsx.
'               Recarga completa.
'
'  Capa 2 - PROCESO (Procesar borra y reconstruye)
'    Eventos    una fila por marca valida ordenada por fondo y Dia, con el peso
'               al ultimo FMS con fecha <= Dia (el FMS del mismo dia aun
'               no refleja la marca; si no hay, el anterior). Fecha, valor
'               y AUM los ubica la macro; peso / contribucion / ln(1+r)
'               son formulas de la fila. Re-marcas: encadenadas, salvo
'               las del mes EEFF de Config!C8 (diciembre, auditada vs no
'               auditada): ahi solo cuenta la ultima; la anterior queda con
'               Var Adj 0 y flag "reemplazada por auditada".
'    Matriz     vista horizontal: fondos x meses, tres bloques (retorno
'               del mes por Dia, contribucion F1 pb, contribucion F2 pb).
'    NAV        vista horizontal de BD_FMS: Val_Total de cada fondo por fecha
'               FMS (bloque F1 y bloque F2 con su AUM). Valores.
'    Mensual    instrumento x mes (por Dia): # marcas, retorno del mes
'               (Var Adj encadenadas), contribucion F1/F2 en pb, peso al
'               inicio del mes; filas de totales PEN/USD/Total con el
'               retorno del bloque (= contribucion / peso inicial).
'
'  Capa 3 - RESUMEN (CrearResumen lo construye; idempotente)
'    C3 = cualquier fecha. Posicion al ultimo FMS <= C3. Retornos por
'    fondo (anios, YTD, 1M/3M/6M/12M) encadenando marcas por Dia.
'    Contribucion MTD/YTD en pb por fondo y en % por moneda.
'
' PRIMERA VEZ: CrearMonitor (crea Config y las hojas) -> revisar Config
'   -> ActualizarTodo (carga FMS y marcas, procesa) -> CrearResumen.
' RUTINA DIARIA: ActualizarTodo y la fecha en Resumen!C3.
'=====================================================================

'--------------------- nombres de hojas ------------------------------
Private Const H_CONFIG As String = "Config"
Private Const H_BDFMS As String = "BD_FMS"
Private Const H_BDMAR As String = "BD_Marcas"
Private Const H_EVT As String = "Eventos"
Private Const H_MES As String = "Mensual"
Private Const H_NAV As String = "NAV"
Private Const H_MAT As String = "Matriz"
Private Const H_RES As String = "Resumen"

'--------------------- Config (celdas) -------------------------------
Private Const CF_CARPETA As String = "C3"
Private Const CF_MARCAS As String = "C4"
Private Const CF_DESDE As String = "C5"
Private Const CF_HOJAS As String = "C6"
Private Const CF_MATRIZ As String = "C7"
Private Const CF_REMARCA As String = "C8"        ' mes EEFF en que la re-marca REEMPLAZA (12=dic, 13=todos, 0=ninguno)
Private Const CF_ULT_FMS As String = "C9"
Private Const CF_ULT_MAR As String = "C10"
Private Const CF_ULT_PROC As String = "C11"
Private Const CF_CAT_HDR As Long = 13
Private Const CF_CAT_INI As Long = 14                 ' 12 filas: Orden|Instrumento|SBS|Moneda
Private Const N_FONDOS As Long = 12
Private Const DESDE_DEFECTO As String = "2025-12-01"
Private Const CARPETA_DEFECTO As String = _
    "\\PFLMVIPR1FS0222\Area de Inversiones\2. Inversiones Alternativas\05. Fondos Tradicionales\FMS\"
Private Const MARCAS_DEFECTO As String = _
    "\\PFLMVIPR1FS0222\Area de Inversiones\2. Inversiones Alternativas\Marcas.xlsx"

'--------------------- layout comun ----------------------------------
Private Const F_HDR As Long = 4
Private Const F_INI As Long = 5
Private Const F_MAX As Long = 50000

'--------------------- BD_FMS: A Fecha|B FondoAFP|C SBS|D Instr|E Moneda|F Orden|G Valor|H AUM|I Archivo|J Cargado
Private Const BF_FECHA As Long = 1, BF_FONDO As Long = 2, BF_SBS As Long = 3, BF_INSTR As Long = 4, BF_MON As Long = 5
Private Const BF_ORDEN As Long = 6, BF_VALOR As Long = 7, BF_AUM As Long = 8, BF_ARCH As Long = 9, BF_CARG As Long = 10
Private Const BF_NCOL As Long = 10

'--------------------- BD_Marcas: A Hoja|B SBS|C Instr|D Orden|E Dia|F EEFF|G VP|H SBS%|I VarAdj|J Valida|K Motivo|L Archivo|M Cargado
Private Const BM_HOJA As Long = 1, BM_SBS As Long = 2, BM_INSTR As Long = 3, BM_ORDEN As Long = 4
Private Const BM_DIA As Long = 5, BM_EEFF As Long = 6, BM_VP As Long = 7, BM_PCT As Long = 8, BM_VA As Long = 9
Private Const BM_VALIDA As Long = 10, BM_MOTIVO As Long = 11, BM_ARCH As Long = 12, BM_CARG As Long = 13
Private Const BM_NCOL As Long = 13

'--------------------- Eventos: A #|B SBS|C Instr|D Mon|E Dia|F EEFF|G VP|H VarAdj|I FechaFMS|J ValF1|K AUMF1|L PesoF1
'                               M ValF2|N AUMF2|O PesoF2|P ContF1|Q ContF2|R MesDia|S AnioDia|T ln(1+r)|U Flags
Private Const EV_NCOL As Long = 21

'--------------------- Mensual: A Mes|B SBS|C Instr|D Mon|E N|F Retorno mes|G CF1 pb|H CF2 pb|I FMS inicio|J PesoF1 ini|K PesoF2 ini
Private Const MS_NCOL As Long = 11

'--------------------- archivos FMS ----------------------------------
Private Const HOJA_CARTERA As String = "Cartera"
Private Const HOJA_ANEXO As String = "Anexo I"
Private Const FMS_FONDO As Long = 1, FMS_SBS As Long = 4, FMS_VAL As Long = 10
Private Const MARCA_AUM As String = "1.1.1"

'--------------------- archivo Marcas.xlsx ----------------------------
Private Const MK_FILA_SBS As Long = 2, MK_FILA_HDR As Long = 3, MK_FILA_INI As Long = 4
Private Const MK_OFF_EEFF As Long = 1, MK_OFF_VP As Long = 2, MK_OFF_PCT As Long = 3, MK_OFF_VA As Long = 4

'--------------------- Resumen ----------------------------------------
Private Const RS_PEN_TIT As Long = 11, RS_PEN_INI As Long = 13   ' titulo / primera fila tabla PEN
Private Const RS_USD_TIT As Long = 21, RS_USD_INI As Long = 23   ' titulo / primera fila tabla USD
Private Const RS_AUX As String = "Z"                             ' columna de auxiliares (invisible)

'=====================================================================
' ENTRADAS PUBLICAS
'=====================================================================
Public Sub CrearMonitor()
    Cfg
    HojaBDFMS
    HojaBDMarcas
    HojaEventos
    HojaMensual
    HojaMatriz
    HojaNAV
    Cfg().Activate
    MsgBox "Hojas creadas. Revisa Config (rutas, fecha inicial, catalogo) y corre ActualizarTodo; " & _
           "luego CrearResumen.", vbInformation, "CrearMonitor"
End Sub

Public Sub ActualizarTodo()
    Dim inf As String, t0 As Single
    t0 = Timer
    inf = CargarFMS_Interno(False) & vbCrLf
    inf = inf & CargarMarcas_Interno() & vbCrLf
    inf = inf & Procesar_Interno()
    FechaReporteHoy
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Debug.Print inf
    MsgBox inf & vbCrLf & vbCrLf & Format$(Timer - t0, "0") & " s. Detalle en Inmediato (Alt+F11, Ctrl+G).", _
           vbInformation, "ActualizarTodo"
End Sub

'--- Resumen!C3 = hoy (la foto se calcula hasta esa fecha) ------------
Private Sub FechaReporteHoy()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(H_RES)
    If Not ws Is Nothing Then ws.Range("C3").Value = Date
    On Error GoTo 0
End Sub

Public Sub CargarFMS()
    Dim inf As String
    inf = CargarFMS_Interno(False)
    Application.Calculate
    Debug.Print inf: MsgBox inf, vbInformation, "CargarFMS"
End Sub

Public Sub CargarFMS_Todo()
    If MsgBox("Borra BD_FMS y recarga todos los FMS desde la fecha de Config!C5. " & _
              "Abre un archivo por dia; puede tardar varios minutos. Continuar?", _
              vbYesNo + vbQuestion, "CargarFMS_Todo") <> vbYes Then Exit Sub
    Dim inf As String
    inf = CargarFMS_Interno(True)
    Application.Calculate
    Debug.Print inf: MsgBox inf, vbInformation, "CargarFMS_Todo"
End Sub

Public Sub CargarMarcas()
    Dim inf As String
    inf = CargarMarcas_Interno()
    Application.Calculate
    Debug.Print inf: MsgBox inf, vbInformation, "CargarMarcas"
End Sub

Public Sub Procesar()
    Dim inf As String
    inf = Procesar_Interno()
    Application.Calculate
    Debug.Print inf: MsgBox inf, vbInformation, "Procesar"
End Sub

'=====================================================================
' CONFIG Y CATALOGO
'=====================================================================
Private Function Cfg() As Worksheet
    Dim ws As Worksheet, i As Long, k As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(H_CONFIG)
    On Error GoTo 0
    If Not ws Is Nothing Then Set Cfg = ws: Exit Function

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = H_CONFIG
    ws.Range("A1").Value = "Config - Monitor Fondos Tradicionales por dia de marca (celdas amarillas editables)"
    ws.Range("A1").Font.Bold = True
    ws.Range("B3").Value = "Carpeta FMS (varias con ;)"
    ws.Range("B4").Value = "Archivo Marcas.xlsx"
    ws.Range("B5").Value = "Cargar FMS desde"
    ws.Range("B6").Value = "Hojas de Marcas (con ;)"
    ws.Range("B7").Value = "Matriz desde (mes)"
    ws.Range("B8").Value = "Re-marca reemplaza si mes EEFF = (12=dic, 13=todos, 0=no)"
    ws.Range("B9").Value = "Ultima carga FMS"
    ws.Range("B10").Value = "Ultima carga Marcas"
    ws.Range("B11").Value = "Ultimo Procesar"
    ws.Range(CF_CARPETA).Value = CARPETA_DEFECTO
    ws.Range(CF_MARCAS).Value = MARCAS_DEFECTO
    ws.Range(CF_DESDE).Value = DateValue(DESDE_DEFECTO): ws.Range(CF_DESDE).NumberFormat = "dd-mmm-yy"
    ws.Range(CF_HOJAS).Value = "1;2"
    ws.Range(CF_MATRIZ).Value = DateSerial(Year(DateValue(DESDE_DEFECTO)) - 2, 1, 1): ws.Range(CF_MATRIZ).NumberFormat = "mmm-yy"
    ws.Range(CF_REMARCA).Value = 12
    ws.Range("C9:C11").NumberFormat = "dd-mmm-yy hh:mm"
    ws.Range("C3:C8").Interior.Color = RGB(255, 242, 204)

    ws.Cells(CF_CAT_HDR, 2).Value = "Orden": ws.Cells(CF_CAT_HDR, 3).Value = "Instrumento"
    ws.Cells(CF_CAT_HDR, 4).Value = "Codigo SBS": ws.Cells(CF_CAT_HDR, 5).Value = "Moneda"
    Encabezado ws, CF_CAT_HDR, 2, 5
    Dim resp As Variant
    resp = Array( _
        Array("HMC - Credito Peru II (PEN)", "4975141HMPSX", "PEN"), _
        Array("Fondo Credicorp Deuda Soles", "4911011FCDS1", "PEN"), _
        Array("Fondo Credicorp Deuda Soles II", "4911011FCDS2", "PEN"), _
        Array("PEPCO II", "4911021CORF2", "PEN"), _
        Array("Fondo Credicorp Deuda Titulizada", "4911011CCDTI", "PEN"), _
        Array("Fondo HMC Credito Peru III (PEN)", "4975141HM3SX", "PEN"), _
        Array("BD Capital - Senior Loans Clase C", "4974252SCBDC", "USD"), _
        Array("HMC - Credito Peru II (USD)", "4975142HMPDX", "USD"), _
        Array("LV FIAFE", "4911102LINAE", "USD"), _
        Array("BD Capital - Senior Loans 2", "4974252S2BDC", "USD"), _
        Array("Fondo HMC Credito Peru III (USD)", "4975142HM3DX", "USD"), _
        Array("Moneda Patria Peru Fixed Income", "4977602MPAFI", "USD"))
    For i = 0 To N_FONDOS - 1
        k = CF_CAT_INI + i
        ws.Cells(k, 2).Value = i + 1
        ws.Cells(k, 3).Value = resp(i)(0): ws.Cells(k, 4).Value = resp(i)(1): ws.Cells(k, 5).Value = resp(i)(2)
    Next i
    ws.Range(ws.Cells(CF_CAT_INI, 3), ws.Cells(CF_CAT_INI + N_FONDOS - 1, 5)).Interior.Color = RGB(255, 242, 204)
    ws.Columns(2).ColumnWidth = 52: ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 14: ws.Columns(5).ColumnWidth = 9
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
    ws.Range("A1").Font.Size = 10
    Set Cfg = ws
End Function

'--- Catalogo: (1..12, 1)=Instrumento (2)=SBS (3)=Moneda (4)=Orden -------
Private Function Catalogo() As Variant
    Dim ws As Worksheet, arr(1 To N_FONDOS, 1 To 4) As Variant, i As Long
    Set ws = Cfg()
    For i = 1 To N_FONDOS
        arr(i, 1) = CStr(ws.Cells(CF_CAT_INI + i - 1, 3).Value)
        arr(i, 2) = UCase$(Trim$(CStr(ws.Cells(CF_CAT_INI + i - 1, 4).Value)))
        arr(i, 3) = UCase$(Trim$(CStr(ws.Cells(CF_CAT_INI + i - 1, 5).Value)))
        arr(i, 4) = i
    Next i
    Catalogo = arr
End Function

Private Function IdxCatalogo(cat As Variant, sbs As String) As Long
    Dim i As Long
    For i = 1 To N_FONDOS
        If cat(i, 2) = UCase$(sbs) Then IdxCatalogo = i: Exit Function
    Next i
End Function

Private Function FechaDesde() As Date
    Dim v As Variant
    v = Cfg().Range(CF_DESDE).Value
    If IsDate(v) Then FechaDesde = CDate(v) Else FechaDesde = DateValue(DESDE_DEFECTO)
End Function

'=====================================================================
' HOJAS DE DATOS
'=====================================================================
Private Function HojaBDFMS() As Worksheet
    Set HojaBDFMS = HojaDatos(H_BDFMS, "BD_FMS - base bruta de posiciones diarias (una fila por Fecha x Fondo AFP x instrumento). Solo escribe la macro.", _
        Array("Fecha", "Fondo AFP", "Codigo SBS", "Instrumento", "Moneda", "Orden", "Valor S/", "AUM Fondo AFP S/", "Archivo", "Cargado"))
End Function

Private Function HojaBDMarcas() As Worksheet
    Set HojaBDMarcas = HojaDatos(H_BDMAR, "BD_Marcas - base bruta de marcas tal como vienen de Marcas.xlsx. Solo escribe la macro (recarga completa).", _
        Array("Hoja", "Codigo SBS", "Instrumento", "Orden", "Día", "EEFF", "VP", "SBS %", "Var Adj", "Válida", "Motivo", "Archivo", "Cargado"))
End Function

Private Function HojaEventos() As Worksheet
    Set HojaEventos = HojaDatos(H_EVT, "Eventos - una fila por marca valida ordenada por Dia; peso al ultimo FMS con fecha <= Dia, que aun no refleja la marca (I-K, M-N los ubica Procesar); peso, contribucion y ln(1+r) en formulas de la fila.", _
        Array("#", "Codigo SBS", "Instrumento", "Moneda", "Día", "EEFF", "VP", "Var Adj", _
              "Fecha FMS", "Valor F1", "AUM F1", "Peso F1", "Valor F2", "AUM F2", "Peso F2", _
              "Contrib F1", "Contrib F2", "Mes Día", "Año Día", "ln(1+VarAdj)", "Flags"))
End Function

Private Function HojaMensual() As Worksheet
    Set HojaMensual = HojaDatos(H_MES, "Mensual - por instrumento y mes (segun Dia de marca): retorno del mes (Var Adj encadenadas), contribucion a F1/F2 en pb y peso al inicio del mes. Totales = retorno del bloque.", _
        Array("Mes", "Codigo SBS", "Instrumento", "Moneda", "N marcas", "Retorno mes", "Contrib F1 (pb)", "Contrib F2 (pb)", "FMS inicio", "Peso F1 inicio", "Peso F2 inicio"))
End Function

Private Function HojaMatriz() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(H_MAT)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = H_MAT
        ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
    End If
    ws.Range("A1").Value = "Matriz - fondos x meses segun Dia de marca: retorno del mes, contribucion a F1 y a F2 (pb). La reconstruye Procesar; el mes inicial esta en Config."
    ws.Range("A1").Font.Bold = True
    Set HojaMatriz = ws
End Function

'=====================================================================
' CAPA 1a: BD_FMS (carga incremental de FMS_AAAAMMDD)
'=====================================================================
Private Function CargarFMS_Interno(todo As Boolean) As String
    Dim ws As Worksheet, cat As Variant, desde As Date, inf As String
    Dim fechas() As String, rutas() As String, n As Long, i As Long
    Dim calcPrev As Long, fila As Long, ok As Long, mal As Long

    calcPrev = Application.Calculation
    On Error GoTo Manejo
    Set ws = HojaBDFMS()
    cat = Catalogo()
    desde = FechaDesde()
    If todo Then LimpiarDatos ws, BF_NCOL

    Dim dExist As Object
    Set dExist = CreateObject("Scripting.Dictionary")
    For i = F_INI To UltimaFila(ws, BF_FECHA)
        If IsDate(ws.Cells(i, BF_FECHA).Value) Then dExist(CLng(CDate(ws.Cells(i, BF_FECHA).Value))) = 1
    Next i

    n = ListarFMS(desde, dExist, fechas, rutas)
    If n = 0 Then
        CargarFMS_Interno = "BD_FMS: sin FMS nuevos desde " & Format$(desde, "dd-mmm-yy") & _
                            " (ultima fecha " & Format$(UltimaFechaFMS(ws), "dd-mmm-yy") & ")."
        Exit Function
    End If

    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    fila = UltimaFila(ws, BF_FECHA) + 1
    For i = 1 To n
        Application.StatusBar = "BD_FMS " & i & "/" & n & "  " & Mid$(rutas(i), InStrRev(rutas(i), "\") + 1)
        If CargarUnFMS(ws, fila, fechas(i), rutas(i), cat, inf) Then ok = ok + 1 Else mal = mal + 1
    Next i
    OrdenarRango ws, BF_NCOL, BF_FECHA, BF_FONDO, BF_ORDEN
    FormatoBDFMS ws
    Cfg().Range(CF_ULT_FMS).Value = Now

Limpieza:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True: Application.ScreenUpdating = True
    CargarFMS_Interno = "BD_FMS: " & ok & " dias cargados" & IIf(mal > 0, ", " & mal & " con error", "") & _
                        ". Ultima fecha: " & Format$(UltimaFechaFMS(ws), "dd-mmm-yy") & vbCrLf & inf
    Exit Function
Manejo:
    inf = "ERROR " & Err.Number & ": " & Err.Description & vbCrLf & inf
    Resume Limpieza
End Function

Private Function UltimaFechaFMS(ws As Worksheet) As Variant
    Dim r As Long
    r = UltimaFila(ws, BF_FECHA)
    If r >= F_INI Then UltimaFechaFMS = ws.Cells(r, BF_FECHA).Value Else UltimaFechaFMS = Empty
End Function

Private Sub FormatoBDFMS(ws As Worksheet)
    Dim ult As Long
    ult = UltimaFila(ws, BF_FECHA)
    If ult < F_INI Then Exit Sub
    With ws
        .Range(.Cells(F_INI, BF_FECHA), .Cells(ult, BF_FECHA)).NumberFormat = "dd-mmm-yy"
        .Range(.Cells(F_INI, BF_VALOR), .Cells(ult, BF_AUM)).NumberFormat = "#,##0"
        .Range(.Cells(F_INI, BF_CARG), .Cells(ult, BF_CARG)).NumberFormat = "dd-mmm-yy hh:mm"
        .Columns(BF_FECHA).ColumnWidth = 10: .Columns(BF_FONDO).ColumnWidth = 9
        .Columns(BF_SBS).ColumnWidth = 13: .Columns(BF_INSTR).ColumnWidth = 32: .Columns(BF_MON).ColumnWidth = 7
        .Columns(BF_ORDEN).ColumnWidth = 6: .Columns(BF_VALOR).ColumnWidth = 14
        .Columns(BF_AUM).ColumnWidth = 16: .Columns(BF_ARCH).ColumnWidth = 20: .Columns(BF_CARG).ColumnWidth = 14
    End With
End Sub

'--- Un archivo FMS -> 24 filas (12 fondos x F1/F2) -------------------
Private Function CargarUnFMS(ws As Worksheet, ByRef fila As Long, f8 As String, ruta As String, _
                             cat As Variant, ByRef inf As String) As Boolean
    Dim wbF As Workbook, wsF As Worksheet, nombre As String, abriYo As Boolean
    Dim numErr As Long, descErr As String, filaIni As Long
    filaIni = fila
    On Error GoTo Fallo
    nombre = Mid$(ruta, InStrRev(ruta, "\") + 1)
    On Error Resume Next
    Set wbF = Workbooks(nombre)
    On Error GoTo Fallo
    If wbF Is Nothing Then
        Set wbF = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0)
        abriYo = True
    End If
    Set wsF = BuscarHoja(wbF, HOJA_CARTERA)
    If wsF Is Nothing Then Set wsF = wbF.Worksheets(1)

    Dim dPos As Object, dAUM As Object, diag As String
    Set dPos = CreateObject("Scripting.Dictionary"): Set dAUM = CreateObject("Scripting.Dictionary")
    LeerFMS wsF, dPos, dAUM, diag
    If Not (dAUM.Exists("01") And dAUM.Exists("02")) Then
        Dim wsA As Worksheet
        Set wsA = BuscarHoja(wbF, HOJA_ANEXO)
        If Not wsA Is Nothing Then LeerAUM wsA, dAUM
    End If
    If Not (dAUM.Exists("01") And dAUM.Exists("02")) Then
        Dim wsX As Worksheet
        For Each wsX In wbF.Worksheets
            If Not wsX Is wsF Then LeerAUM wsX, dAUM
            If dAUM.Exists("01") And dAUM.Exists("02") Then Exit For
        Next wsX
    End If
    If abriYo Then wbF.Close SaveChanges:=False

    Dim fecha As Date, i As Long, f As Long, fondo As String, k As Long, faltan As String, clave As String
    fecha = DateSerial(CLng(Left$(f8, 4)), CLng(Mid$(f8, 5, 2)), CLng(Right$(f8, 2)))
    For f = 1 To 2
        fondo = Format$(f, "00")
        For i = 1 To N_FONDOS
            clave = fondo & "|" & cat(i, 2)
            ws.Cells(fila, BF_FECHA).Value = fecha
            ws.Cells(fila, BF_FONDO).Value = "F" & f
            ws.Cells(fila, BF_SBS).Value = cat(i, 2)
            ws.Cells(fila, BF_INSTR).Value = cat(i, 1)
            ws.Cells(fila, BF_MON).Value = cat(i, 3)
            ws.Cells(fila, BF_ORDEN).Value = i
            If dPos.Exists(clave) Then
                ws.Cells(fila, BF_VALOR).Value = dPos(clave): k = k + 1
            Else
                ws.Cells(fila, BF_VALOR).Value = 0
                faltan = faltan & IIf(faltan = "", "", ", ") & cat(i, 1) & " (F" & f & ")"
            End If
            If dAUM.Exists(fondo) Then ws.Cells(fila, BF_AUM).Value = dAUM(fondo)
            ws.Cells(fila, BF_ARCH).Value = nombre
            ws.Cells(fila, BF_CARG).Value = Now
            fila = fila + 1
        Next i
    Next f
    inf = inf & nombre & " -> " & Format$(fecha, "dd-mmm-yy") & ": " & k & " posiciones [" & diag & "]" & _
          IIf(dAUM.Exists("01"), ", AUM F1 ok", ", AUM F1 NO ENCONTRADO") & _
          IIf(dAUM.Exists("02"), ", AUM F2 ok", ", AUM F2 NO ENCONTRADO") & _
          IIf(k = 0, " | NO LEI POSICIONES", IIf(faltan <> "", " | sin posicion: " & faltan, "")) & vbCrLf
    CargarUnFMS = True
    Exit Function
Fallo:
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    If abriYo And Not wbF Is Nothing Then wbF.Close SaveChanges:=False
    If fila > filaIni Then ws.Range(ws.Cells(filaIni, 1), ws.Cells(fila, BF_NCOL)).ClearContents
    fila = filaIni
    inf = inf & nombre & " -> ERROR " & numErr & ": " & descErr & " (saltado)" & vbCrLf
End Function

'=====================================================================
' CAPA 1b: BD_Marcas (recarga completa desde Marcas.xlsx)
'=====================================================================
Private Function CargarMarcas_Interno() As String
    Dim ws As Worksheet, cat As Variant, wbM As Workbook, abriYo As Boolean, inf As String
    Dim calcPrev As Long, ruta As String, nombre As String, hojas() As String, h As Long
    Dim wsM As Worksheet, i As Long, col As Long, fila As Long, k As Long, tot As Long

    calcPrev = Application.Calculation
    On Error GoTo Manejo
    Set ws = HojaBDMarcas()
    cat = Catalogo()
    ruta = Trim$(CStr(Cfg().Range(CF_MARCAS).Value))
    nombre = Mid$(ruta, InStrRev(ruta, "\") + 1)
    On Error Resume Next
    Set wbM = Workbooks(nombre)
    On Error GoTo Manejo
    If wbM Is Nothing Then
        If Dir$(ruta) = vbNullString Then
            CargarMarcas_Interno = "BD_Marcas: no encuentro " & ruta & " (Config!C4)."
            Exit Function
        End If
        Set wbM = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0)
        abriYo = True
    End If

    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    LimpiarDatos ws, BM_NCOL
    fila = F_INI
    hojas = Split(CStr(Cfg().Range(CF_HOJAS).Value), ";")
    For h = LBound(hojas) To UBound(hojas)
        Set wsM = Nothing
        On Error Resume Next
        Set wsM = wbM.Worksheets(Trim$(hojas(h)))
        On Error GoTo Manejo
        If Not wsM Is Nothing Then
            Dim ultFila As Long, ultCol As Long
            UsadoReal wsM, ultFila, ultCol
            For i = 1 To N_FONDOS
                col = BuscarCodigo(wsM, CStr(cat(i, 2)), ultCol)
                If col > 0 Then
                    k = LeerBloqueMarcas(wsM, col, i, cat, ultFila, ws, fila, nombre)
                    tot = tot + k
                    inf = inf & "Hoja " & wsM.Name & " | " & cat(i, 1) & ": " & k & " marcas (col " & _
                          Split(wsM.Cells(1, col).Address, "$")(1) & ")" & vbCrLf
                End If
            Next i
        End If
    Next h
    If abriYo Then wbM.Close SaveChanges:=False
    OrdenarRango ws, BM_NCOL, BM_DIA, BM_ORDEN, 0
    FormatoBDMarcas ws
    Cfg().Range(CF_ULT_MAR).Value = Now

Limpieza:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True: Application.ScreenUpdating = True
    CargarMarcas_Interno = "BD_Marcas: " & tot & " marcas cargadas de " & nombre & "." & vbCrLf & inf
    Exit Function
Manejo:
    inf = "ERROR " & Err.Number & ": " & Err.Description & vbCrLf & inf
    On Error Resume Next
    If abriYo And Not wbM Is Nothing Then wbM.Close SaveChanges:=False
    Resume Limpieza
End Function

'=====================================================================
' CAPA 2: PROCESAR -> Eventos, Mensual
'=====================================================================
Private Function Procesar_Interno() As String
    Dim cat As Variant, wsB As Worksheet, wsM As Worksheet, wsE As Worksheet, inf As String
    Dim calcPrev As Long
    calcPrev = Application.Calculation
    On Error GoTo Manejo
    cat = Catalogo()
    Set wsB = HojaBDFMS()
    Set wsM = HojaBDMarcas()
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ---- 1. indice de BD_FMS: fechas distintas + valores por clave ----
    Dim dVal As Object, dAUM As Object, dFec As Object, r As Long, ult As Long, key As String
    Set dVal = CreateObject("Scripting.Dictionary"): Set dAUM = CreateObject("Scripting.Dictionary")
    Set dFec = CreateObject("Scripting.Dictionary")
    ult = UltimaFila(wsB, BF_FECHA)
    Dim datosB As Variant
    If ult >= F_INI Then datosB = wsB.Range(wsB.Cells(F_INI, 1), wsB.Cells(ult, BF_NCOL)).Value
    Dim fLng As Long, fondo As String
    For r = 1 To ult - F_INI + 1
        If IsDate(datosB(r, BF_FECHA)) Then
            fLng = CLng(CDate(datosB(r, BF_FECHA)))
            fondo = UCase$(CStr(datosB(r, BF_FONDO)))
            dFec(fLng) = 1
            key = fLng & "|" & fondo & "|" & UCase$(CStr(datosB(r, BF_SBS)))
            If EsNumero(datosB(r, BF_VALOR)) Then dVal(key) = CDbl(datosB(r, BF_VALOR)) Else dVal(key) = 0
            If EsNumero(datosB(r, BF_AUM)) Then dAUM(fLng & "|" & fondo) = CDbl(datosB(r, BF_AUM))
        End If
    Next r
    Dim fechasB() As Long, nF As Long, kk As Variant, i As Long, j As Long, t As Long
    nF = dFec.Count
    If nF > 0 Then
        ReDim fechasB(1 To nF)
        For Each kk In dFec.Keys
            i = i + 1: fechasB(i) = CLng(kk)
        Next kk
        For i = 1 To nF - 1
            For j = i + 1 To nF
                If fechasB(j) < fechasB(i) Then t = fechasB(i): fechasB(i) = fechasB(j): fechasB(j) = t
            Next j
        Next i
    End If

    ' ---- 2. marcas validas, ordenadas por Dia (estable) ----
    ult = UltimaFila(wsM, BM_DIA)
    If ult < F_INI Then
        Procesar_Interno = "Procesar: BD_Marcas esta vacia; corre CargarMarcas primero."
        GoTo Limpieza
    End If
    Dim datosM As Variant, nM As Long, idx() As Long, n As Long
    datosM = wsM.Range(wsM.Cells(F_INI, 1), wsM.Cells(ult, BM_NCOL)).Value
    nM = ult - F_INI + 1
    ReDim idx(1 To nM)
    For r = 1 To nM
        If UCase$(CStr(datosM(r, BM_VALIDA))) = "SI" And IsDate(datosM(r, BM_DIA)) Then
            n = n + 1: idx(n) = r
        End If
    Next r
    If n = 0 Then
        Procesar_Interno = "Procesar: no hay marcas validas en BD_Marcas."
        GoTo Limpieza
    End If
    For i = 2 To n                                    ' orden: fondo (catalogo) y luego Dia; estable
        t = idx(i): j = i - 1
        Do While j >= 1
            If ClaveOrden(datosM, idx(j)) <= ClaveOrden(datosM, t) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = t
    Next i

    ' ---- 3. Eventos ----
    ' re-marcas que REEMPLAZAN (auditada vs no auditada): ultima marca por (fondo, EEFF) en el mes de Config!C8
    Dim mesRegla As Long, dUltima As Object, pK As String, p As Long
    If Trim$(Cfg().Range(CF_REMARCA).Value & "") = vbNullString Then     ' Config creada por una version anterior
        Cfg().Range("B8").Value = "Re-marca reemplaza si mes EEFF = (12=dic, 13=todos, 0=no)"
        Cfg().Range(CF_REMARCA).Value = 12
        Cfg().Range(CF_REMARCA).Interior.Color = RGB(255, 242, 204)
    End If
    mesRegla = Val(Cfg().Range(CF_REMARCA).Value & "")
    Set dUltima = CreateObject("Scripting.Dictionary")
    For p = 1 To n
        r = idx(p)
        If IsDate(datosM(r, BM_EEFF)) Then
            If mesRegla = 13 Or Month(CDate(datosM(r, BM_EEFF))) = mesRegla Then
                pK = UCase$(CStr(datosM(r, BM_SBS))) & "|" & CLng(CDate(datosM(r, BM_EEFF)))
                dUltima(pK) = p                                   ' se queda con la ultima (orden fondo/Dia/lectura)
            End If
        End If
    Next p
    Set wsE = HojaEventos()
    LimpiarDatos wsE, EV_NCOL
    Dim salida() As Variant, dSeen As Object, ic As Long, dia As Date, eeff As Variant, va As Double
    Dim sbs As String, fT As Long, flags As String, encontrado As Boolean, diaAnt As Object
    ReDim salida(1 To n, 1 To EV_NCOL)
    Set dSeen = CreateObject("Scripting.Dictionary"): Set diaAnt = CreateObject("Scripting.Dictionary")
    Dim sinPeso As Long, sinFMS As Long, remarcas As Long, reemplazadas As Long, reemplazada As Boolean
    For p = 1 To n
        r = idx(p)
        sbs = UCase$(CStr(datosM(r, BM_SBS))): ic = IdxCatalogo(cat, sbs)
        dia = CDate(datosM(r, BM_DIA)): eeff = datosM(r, BM_EEFF): va = CDbl(datosM(r, BM_VA))
        flags = vbNullString
        reemplazada = False
        If IsDate(eeff) Then
            pK = sbs & "|" & CLng(CDate(eeff))
            If dUltima.Exists(pK) Then
                If dUltima(pK) <> p Then
                    reemplazada = True: reemplazadas = reemplazadas + 1
                    va = 0: flags = "reemplazada por auditada"
                End If
            End If
        End If
        salida(p, 1) = p: salida(p, 2) = sbs: salida(p, 3) = cat(ic, 1): salida(p, 4) = cat(ic, 3)
        salida(p, 5) = dia
        If IsDate(eeff) Then salida(p, 6) = CDate(eeff) Else salida(p, 6) = Empty
        salida(p, 7) = datosM(r, BM_VP): salida(p, 8) = va
        encontrado = False
        For i = nF To 1 Step -1
            If fechasB(i) <= CLng(dia) Then fT = fechasB(i): encontrado = True: Exit For   ' el FMS del mismo dia aun no trae la marca
        Next i
        If encontrado Then
            salida(p, 9) = CDate(fT)
            salida(p, 10) = ValorDe(dVal, fT & "|F1|" & sbs): salida(p, 11) = ValorDe(dAUM, fT & "|F1")
            salida(p, 13) = ValorDe(dVal, fT & "|F2|" & sbs): salida(p, 14) = ValorDe(dAUM, fT & "|F2")
            If salida(p, 10) = 0 And salida(p, 13) = 0 Then
                flags = "sin posicion": sinPeso = sinPeso + 1
            End If
        Else
            flags = "sin FMS": sinFMS = sinFMS + 1
        End If
        If IsDate(eeff) Then
            key = sbs & "|" & CLng(CDate(eeff))
            If dSeen.Exists(key) Then
                flags = flags & IIf(flags = "", "", "; ") & "re-marca": remarcas = remarcas + 1
            Else
                dSeen(key) = 1
            End If
            If CDate(eeff) <> DateSerial(Year(CDate(eeff)), Month(CDate(eeff)) + 1, 0) Then
                flags = flags & IIf(flags = "", "", "; ") & "EEFF no es fin de mes"
            End If
        End If
        If va = 0 And Not reemplazada Then flags = flags & IIf(flags = "", "", "; ") & "Var Adj = 0"
        salida(p, 18) = DateSerial(Year(dia), Month(dia), 1): salida(p, 19) = Year(dia)
        salida(p, 21) = flags
    Next p
    wsE.Range(wsE.Cells(F_INI, 1), wsE.Cells(F_INI + n - 1, EV_NCOL)).Value = salida
    Dim r0 As String, ultE As Long
    r0 = CStr(F_INI): ultE = F_INI + n - 1
    wsE.Range(wsE.Cells(F_INI, 12), wsE.Cells(ultE, 12)).Formula = "=IF(OR($J" & r0 & "="""",$K" & r0 & "=""""),"""",IFERROR($J" & r0 & "/$K" & r0 & ",""""))"
    wsE.Range(wsE.Cells(F_INI, 15), wsE.Cells(ultE, 15)).Formula = "=IF(OR($M" & r0 & "="""",$N" & r0 & "=""""),"""",IFERROR($M" & r0 & "/$N" & r0 & ",""""))"
    wsE.Range(wsE.Cells(F_INI, 16), wsE.Cells(ultE, 16)).Formula = "=IF($L" & r0 & "="""","""",$L" & r0 & "*$H" & r0 & ")"
    wsE.Range(wsE.Cells(F_INI, 17), wsE.Cells(ultE, 17)).Formula = "=IF($O" & r0 & "="""","""",$O" & r0 & "*$H" & r0 & ")"
    wsE.Range(wsE.Cells(F_INI, 20), wsE.Cells(ultE, 20)).Formula = "=IFERROR(LN(1+$H" & r0 & "),"""")"
    FormatoEventos wsE, ultE

    ' ---- 4. Mensual ----
    Dim mesMin As Date, mesMax As Date
    mesMin = CDate(salida(1, 18)): mesMax = CDate(salida(1, 18))
    For p = 1 To n
        If CDate(salida(p, 18)) < mesMin Then mesMin = CDate(salida(p, 18))
        If CDate(salida(p, 18)) > mesMax Then mesMax = CDate(salida(p, 18))
    Next p
    EscribirMensual cat, mesMin, mesMax
    EscribirMatriz cat, mesMax
    EscribirNAV cat, fechasB, nF, dVal, dAUM

    Cfg().Range(CF_ULT_PROC).Value = Now
    inf = "Procesar: " & n & " eventos (" & remarcas & " re-marcas, de ellas " & reemplazadas & " reemplazadas por auditada; " & _
          sinFMS & " sin FMS, " & sinPeso & " sin posicion). Mensual " & Format$(mesMin, "mmm-yy") & " a " & Format$(mesMax, "mmm-yy") & "."

Limpieza:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True: Application.ScreenUpdating = True
    If Procesar_Interno = vbNullString Then Procesar_Interno = inf
    Exit Function
Manejo:
    inf = "ERROR " & Err.Number & ": " & Err.Description & " | " & inf
    Resume Limpieza
End Function

Private Function ClaveOrden(datosM As Variant, r As Long) As Double
    ' orden del catalogo * 1e6 + fecha (serial): agrupa por fondo y ordena por Dia
    ClaveOrden = CDbl(Val(datosM(r, BM_ORDEN) & "")) * 1000000# + CDbl(CLng(CDate(datosM(r, BM_DIA))))
End Function

Private Function ValorDe(d As Object, key As String) As Double
    If d.Exists(key) Then ValorDe = CDbl(d(key)) Else ValorDe = 0
End Function

Private Sub FormatoEventos(ws As Worksheet, ultE As Long)
    With ws
        .Range(.Cells(F_INI, 5), .Cells(ultE, 6)).NumberFormat = "dd-mmm-yy"
        .Range(.Cells(F_INI, 7), .Cells(ultE, 7)).NumberFormat = "#,##0.0000"
        .Range(.Cells(F_INI, 8), .Cells(ultE, 8)).NumberFormat = "0.000%"
        .Range(.Cells(F_INI, 9), .Cells(ultE, 9)).NumberFormat = "dd-mmm-yy"
        .Range(.Cells(F_INI, 10), .Cells(ultE, 11)).NumberFormat = "#,##0"
        .Range(.Cells(F_INI, 12), .Cells(ultE, 12)).NumberFormat = "0.000%"
        .Range(.Cells(F_INI, 13), .Cells(ultE, 14)).NumberFormat = "#,##0"
        .Range(.Cells(F_INI, 15), .Cells(ultE, 15)).NumberFormat = "0.000%"
        .Range(.Cells(F_INI, 16), .Cells(ultE, 17)).NumberFormat = "0.0000%"
        .Range(.Cells(F_INI, 18), .Cells(ultE, 18)).NumberFormat = "mmm-yy"
        .Range(.Cells(F_INI, 20), .Cells(ultE, 20)).NumberFormat = "0.00000"
        .Columns(1).ColumnWidth = 5: .Columns(2).ColumnWidth = 13: .Columns(3).ColumnWidth = 32
        .Columns(4).ColumnWidth = 7
        Dim c As Long
        For c = 5 To EV_NCOL: .Columns(c).ColumnWidth = 11: Next c
        .Columns(10).ColumnWidth = 13: .Columns(11).ColumnWidth = 15
        .Columns(13).ColumnWidth = 13: .Columns(14).ColumnWidth = 15
        .Columns(EV_NCOL).ColumnWidth = 30
    End With
End Sub

'--- Matriz: fondos x meses (retorno, contrib F1 pb, contrib F2 pb) -------
Private Sub EscribirMatriz(cat As Variant, mesMax As Date)
    Dim ws As Worksheet, mesIni As Date, m As Date, nMes As Long, c As Long, i As Long, k As Long, b As Long, r As Long, r0 As Long
    Dim EQ As String, BQ As String, key As String, mCel As String
    Set ws = HojaMatriz()
    ws.Range("A2:ZZ200").Clear
    ws.Range("A2:ZZ200").Font.Name = "Arial": ws.Range("A2:ZZ200").Font.Size = 8
    Dim v As Variant: v = Cfg().Range(CF_MATRIZ).Value
    If Not IsDate(v) Then                                             ' Config creada por una version anterior
        Cfg().Range("B7").Value = "Matriz desde (mes)"
        Cfg().Range(CF_MATRIZ).Value = DateSerial(Year(FechaDesde()) - 2, 1, 1): Cfg().Range(CF_MATRIZ).NumberFormat = "mmm-yy"
        Cfg().Range(CF_MATRIZ).Interior.Color = RGB(255, 242, 204)
        v = Cfg().Range(CF_MATRIZ).Value
    End If
    If IsDate(v) Then mesIni = DateSerial(Year(CDate(v)), Month(CDate(v)), 1) Else mesIni = DateSerial(Year(FechaDesde()) - 2, 1, 1)
    If mesIni > mesMax Then mesIni = DateSerial(Year(mesMax), 1, 1)
    EQ = H_EVT & "!": BQ = H_BDFMS & "!"
    Dim eS As String, eD As String, eMo As String, eM As String, eP As String, eCF2 As String, eL As String
    eS = EQ & "$B$" & F_INI & ":$B$" & F_MAX: eD = EQ & "$E$" & F_INI & ":$E$" & F_MAX: eMo = EQ & "$D$" & F_INI & ":$D$" & F_MAX
    eM = EQ & "$R$" & F_INI & ":$R$" & F_MAX: eP = EQ & "$P$" & F_INI & ":$P$" & F_MAX
    eCF2 = EQ & "$Q$" & F_INI & ":$Q$" & F_MAX: eL = EQ & "$T$" & F_INI & ":$T$" & F_MAX
    Dim bF As String: bF = BQ & "$A$" & F_INI & ":$A$" & F_MAX
    Dim titulos As Variant: titulos = Array("RETORNO DEL MES (por Día de marca)", "CONTRIBUCIÓN A FONDO 1 (pb)", "CONTRIBUCIÓN A FONDO 2 (pb)")
    Dim mon As Variant: mon = Array("PEN", "USD", "Total")
    Dim monLbl As Variant: monLbl = Array("Total PEN", "Total USD", "Total Tradicionales")
    nMes = (Year(mesMax) - Year(mesIni)) * 12 + Month(mesMax) - Month(mesIni) + 1
    For b = 0 To 2
        r0 = 3 + b * 19                                                   ' titulo del bloque
        ws.Cells(r0, 2).Value = titulos(b): ws.Cells(r0, 2).Font.Bold = True
        ws.Cells(r0 + 1, 2).Value = "Instrumento"
        m = mesIni
        For c = 1 To nMes
            ws.Cells(r0 + 1, 2 + c).Value = m: ws.Cells(r0 + 1, 2 + c).NumberFormat = "mmm-yy"
            ws.Cells(r0 + 1, 2 + c).HorizontalAlignment = xlRight
            m = DateSerial(Year(m), Month(m) + 1, 1)
        Next c
        Encabezado ws, r0 + 1, 2, 2 + nMes
        For i = 1 To N_FONDOS
            r = r0 + 1 + i
            ws.Cells(r, 2).Value = cat(i, 1)
            key = H_CONFIG & "!$D$" & (CF_CAT_INI + i - 1)
            For c = 1 To nMes
                mCel = ColLetra(2 + c) & "$" & (r0 + 1)
                Select Case b
                    Case 0
                        ws.Cells(r, 2 + c).Formula = "=IF(COUNTIFS(" & eS & "," & key & "," & eM & "," & mCel & ")=0,""""," & _
                            "EXP(SUMIFS(" & eL & "," & eS & "," & key & "," & eM & "," & mCel & "))-1)"
                    Case 1      ' solo marcas con peso (contribucion numerica)
                        ws.Cells(r, 2 + c).Formula = "=IF(COUNTIFS(" & eS & "," & key & "," & eM & "," & mCel & "," & eP & ","">=-1"")=0,""""," & _
                            "SUMIFS(" & eP & "," & eS & "," & key & "," & eM & "," & mCel & ")*10000)"
                    Case 2
                        ws.Cells(r, 2 + c).Formula = "=IF(COUNTIFS(" & eS & "," & key & "," & eM & "," & mCel & "," & eCF2 & ","">=-1"")=0,""""," & _
                            "SUMIFS(" & eCF2 & "," & eS & "," & key & "," & eM & "," & mCel & ")*10000)"
                End Select
            Next c
        Next i
        For k = 0 To 2                                                    ' totales
            r = r0 + 1 + N_FONDOS + 1 + k
            ws.Cells(r, 2).Value = monLbl(k): ws.Cells(r, 2).Font.Bold = True
            For c = 1 To nMes
                mCel = ColLetra(2 + c) & "$" & (r0 + 1)
                Dim crit As String: crit = IIf(k < 2, "," & eMo & ",""" & mon(k) & """", "")
                Select Case b
                    Case 0      ' retorno del bloque = contribucion F1 del mes / peso F1 del bloque al FMS anterior al mes
                        ws.Cells(r, 2 + c).Formula = "=IF(COUNTIFS(" & eM & "," & mCel & crit & ")=0,"""",IFERROR(SUMIFS(" & eP & "," & eM & "," & mCel & crit & ")/" & _
                            PesoBloque("F1", CStr(mon(k)), "MAXIFS(" & bF & "," & bF & ",""<""&" & mCel & ")", BQ) & ",""""))"
                    Case 1
                        ws.Cells(r, 2 + c).Formula = "=IF(COUNTIFS(" & eM & "," & mCel & crit & "," & eP & ","">=-1"")=0,"""",SUMIFS(" & eP & "," & eM & "," & mCel & crit & ")*10000)"
                    Case 2
                        ws.Cells(r, 2 + c).Formula = "=IF(COUNTIFS(" & eM & "," & mCel & crit & "," & eCF2 & ","">=-1"")=0,"""",SUMIFS(" & eCF2 & "," & eM & "," & mCel & crit & ")*10000)"
                End Select
            Next c
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 2 + nMes)).Font.Bold = True
        Next k
        With ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r0 + 1 + N_FONDOS + 4, 2 + nMes))
            .NumberFormat = IIf(b = 0, "0.00%", "#,##0.0")
        End With
        Heatmap ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r0 + 1 + N_FONDOS, 2 + nMes))
    Next b
    ws.Columns(1).ColumnWidth = 2: ws.Columns(2).ColumnWidth = 34
    For c = 3 To 2 + nMes: ws.Columns(c).ColumnWidth = 8: Next c
    ws.Activate
    ws.Range("C5").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

'--- NAV: Val_Total por fondo y fecha FMS, horizontal (valores desde BD_FMS) --
Private Function HojaNAV() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(H_NAV)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = H_NAV
        ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
    End If
    ws.Range("A1").Value = "NAV - Val_Total (S/) de cada fondo por fecha FMS, tomado de BD_FMS. Bloque Fondo 1 y bloque Fondo 2 con su AUM. Lo reconstruye Procesar."
    ws.Range("A1").Font.Bold = True
    Set HojaNAV = ws
End Function

Private Sub EscribirNAV(cat As Variant, fechasB() As Long, nF As Long, dVal As Object, dAUM As Object)
    Dim ws As Worksheet, b As Long, r0 As Long, i As Long, c As Long, fondo As String, r As Long
    Set ws = HojaNAV()
    With ws.Range(ws.Cells(2, 1), ws.Cells(60, ws.Columns.Count))
        .Clear
        .Font.Name = "Arial": .Font.Size = 8
    End With
    If nF = 0 Then Exit Sub
    Dim salida() As Variant
    For b = 1 To 2
        fondo = "F" & b
        r0 = 3 + (b - 1) * (N_FONDOS + 5)                                 ' 3 y 20
        ws.Cells(r0, 2).Value = "VALOR EN FONDO " & b & " (S/)": ws.Cells(r0, 2).Font.Bold = True
        ws.Cells(r0 + 1, 2).Value = "Instrumento"
        ReDim salida(1 To N_FONDOS + 1, 1 To nF)
        For c = 1 To nF
            ws.Cells(r0 + 1, 2 + c).Value = CDate(fechasB(c))
            For i = 1 To N_FONDOS
                salida(i, c) = ValorDe(dVal, fechasB(c) & "|" & fondo & "|" & cat(i, 2))
            Next i
            salida(N_FONDOS + 1, c) = ValorDe(dAUM, fechasB(c) & "|" & fondo)
        Next c
        For i = 1 To N_FONDOS
            ws.Cells(r0 + 1 + i, 2).Value = cat(i, 1)
        Next i
        ws.Cells(r0 + 2 + N_FONDOS, 2).Value = "AUM Fondo " & b
        ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r0 + 2 + N_FONDOS, 2 + nF)).Value = salida
        ws.Range(ws.Cells(r0 + 1, 3), ws.Cells(r0 + 1, 2 + nF)).NumberFormat = "dd-mmm-yy"
        ws.Range(ws.Cells(r0 + 1, 3), ws.Cells(r0 + 1, 2 + nF)).HorizontalAlignment = xlRight
        Encabezado ws, r0 + 1, 2, 2 + nF
        ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r0 + 2 + N_FONDOS, 2 + nF)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(r0 + 2 + N_FONDOS, 2), ws.Cells(r0 + 2 + N_FONDOS, 2 + nF)).Font.Bold = True
    Next b
    ws.Columns(1).ColumnWidth = 2: ws.Columns(2).ColumnWidth = 34
    For c = 3 To 2 + nF: ws.Columns(c).ColumnWidth = 11: Next c
    ws.Activate
    ws.Range("C5").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

'--- Mensual: instrumento x mes + totales (formulas sobre Eventos / BD_FMS) --
Private Sub EscribirMensual(cat As Variant, mesMin As Date, mesMax As Date)
    Dim ws As Worksheet, fila As Long, m As Date, i As Long, k As Long
    Set ws = HojaMensual()
    LimpiarDatos ws, MS_NCOL
    Dim EQ As String, BQ As String
    EQ = H_EVT & "!": BQ = H_BDFMS & "!"
    Dim eS As String, eD As String, eM As String, eP As String, eCF2 As String, eL As String
    eS = EQ & "$B$" & F_INI & ":$B$" & F_MAX: eD = EQ & "$D$" & F_INI & ":$D$" & F_MAX
    eM = EQ & "$R$" & F_INI & ":$R$" & F_MAX: eP = EQ & "$P$" & F_INI & ":$P$" & F_MAX
    eCF2 = EQ & "$Q$" & F_INI & ":$Q$" & F_MAX: eL = EQ & "$T$" & F_INI & ":$T$" & F_MAX
    Dim bF As String, bFo As String, bS As String, bMo As String, bV As String, bA As String
    bF = BQ & "$A$" & F_INI & ":$A$" & F_MAX: bFo = BQ & "$B$" & F_INI & ":$B$" & F_MAX
    bS = BQ & "$C$" & F_INI & ":$C$" & F_MAX: bMo = BQ & "$E$" & F_INI & ":$E$" & F_MAX
    bV = BQ & "$G$" & F_INI & ":$G$" & F_MAX: bA = BQ & "$H$" & F_INI & ":$H$" & F_MAX
    Dim totLbl As Variant, totMon As Variant
    totLbl = Array("Total PEN", "Total USD", "Total Tradicionales"): totMon = Array("PEN", "USD", "")
    fila = F_INI: m = mesMin
    Do While m <= mesMax
        For i = 1 To N_FONDOS
            ws.Cells(fila, 1).Value = m
            ws.Cells(fila, 2).Value = cat(i, 2): ws.Cells(fila, 3).Value = cat(i, 1): ws.Cells(fila, 4).Value = cat(i, 3)
            ws.Cells(fila, 5).Formula = "=COUNTIFS(" & eS & ",$B" & fila & "," & eM & ",$A" & fila & ")"
            ws.Cells(fila, 6).Formula = "=IF($E" & fila & "=0,"""",EXP(SUMIFS(" & eL & "," & eS & ",$B" & fila & "," & eM & ",$A" & fila & "))-1)"
            ws.Cells(fila, 7).Formula = "=SUMIFS(" & eP & "," & eS & ",$B" & fila & "," & eM & ",$A" & fila & ")*10000"
            ws.Cells(fila, 8).Formula = "=SUMIFS(" & eCF2 & "," & eS & ",$B" & fila & "," & eM & ",$A" & fila & ")*10000"
            ws.Cells(fila, 9).Formula = "=IFERROR(1/(1/MAXIFS(" & bF & "," & bF & ",""<""&$A" & fila & ")),"""")"
            ws.Cells(fila, 10).Formula = "=IF($I" & fila & "="""","""",IFERROR(SUMIFS(" & bV & "," & bF & ",$I" & fila & "," & bFo & ",""F1""," & bS & ",$B" & fila & _
                                         ")/AVERAGEIFS(" & bA & "," & bF & ",$I" & fila & "," & bFo & ",""F1""),""""))"
            ws.Cells(fila, 11).Formula = "=IF($I" & fila & "="""","""",IFERROR(SUMIFS(" & bV & "," & bF & ",$I" & fila & "," & bFo & ",""F2""," & bS & ",$B" & fila & _
                                         ")/AVERAGEIFS(" & bA & "," & bF & ",$I" & fila & "," & bFo & ",""F2""),""""))"
            fila = fila + 1
        Next i
        For k = 0 To 2
            ws.Cells(fila, 1).Value = m
            ws.Cells(fila, 2).Value = UCase$(totLbl(k)): ws.Cells(fila, 3).Value = totLbl(k): ws.Cells(fila, 4).Value = totMon(k)
            If k < 2 Then
                ws.Cells(fila, 5).Formula = "=COUNTIFS(" & eD & ",$D" & fila & "," & eM & ",$A" & fila & ")"
                ws.Cells(fila, 7).Formula = "=SUMIFS(" & eP & "," & eD & ",$D" & fila & "," & eM & ",$A" & fila & ")*10000"
                ws.Cells(fila, 8).Formula = "=SUMIFS(" & eCF2 & "," & eD & ",$D" & fila & "," & eM & ",$A" & fila & ")*10000"
                ws.Cells(fila, 10).Formula = "=IF($I" & fila & "="""","""",IFERROR(SUMIFS(" & bV & "," & bF & ",$I" & fila & "," & bFo & ",""F1""," & bMo & ",$D" & fila & _
                                             ")/AVERAGEIFS(" & bA & "," & bF & ",$I" & fila & "," & bFo & ",""F1""),""""))"
                ws.Cells(fila, 11).Formula = "=IF($I" & fila & "="""","""",IFERROR(SUMIFS(" & bV & "," & bF & ",$I" & fila & "," & bFo & ",""F2""," & bMo & ",$D" & fila & _
                                             ")/AVERAGEIFS(" & bA & "," & bF & ",$I" & fila & "," & bFo & ",""F2""),""""))"
            Else
                ws.Cells(fila, 5).Formula = "=COUNTIFS(" & eM & ",$A" & fila & ")"
                ws.Cells(fila, 7).Formula = "=SUMIFS(" & eP & "," & eM & ",$A" & fila & ")*10000"
                ws.Cells(fila, 8).Formula = "=SUMIFS(" & eCF2 & "," & eM & ",$A" & fila & ")*10000"
                ws.Cells(fila, 10).Formula = "=IF($I" & fila & "="""","""",IFERROR(SUMIFS(" & bV & "," & bF & ",$I" & fila & "," & bFo & ",""F1"")" & _
                                             "/AVERAGEIFS(" & bA & "," & bF & ",$I" & fila & "," & bFo & ",""F1""),""""))"
                ws.Cells(fila, 11).Formula = "=IF($I" & fila & "="""","""",IFERROR(SUMIFS(" & bV & "," & bF & ",$I" & fila & "," & bFo & ",""F2"")" & _
                                             "/AVERAGEIFS(" & bA & "," & bF & ",$I" & fila & "," & bFo & ",""F2""),""""))"
            End If
            ' retorno del bloque = contribucion F1 / peso inicial F1 (en % del fondo)
            ws.Cells(fila, 6).Formula = "=IF(OR($J" & fila & "="""",$J" & fila & "=0),"""",$G" & fila & "/10000/$J" & fila & ")"
            ws.Cells(fila, 9).Formula = "=IFERROR(1/(1/MAXIFS(" & bF & "," & bF & ",""<""&$A" & fila & ")),"""")"
            ws.Range(ws.Cells(fila, 1), ws.Cells(fila, MS_NCOL)).Font.Bold = True
            fila = fila + 1
        Next k
        m = DateSerial(Year(m), Month(m) + 1, 1)
    Loop
    Dim ult As Long: ult = fila - 1
    With ws
        .Range(.Cells(F_INI, 1), .Cells(ult, 1)).NumberFormat = "mmm-yy"
        .Range(.Cells(F_INI, 6), .Cells(ult, 6)).NumberFormat = "0.00%"
        .Range(.Cells(F_INI, 7), .Cells(ult, 8)).NumberFormat = "#,##0.0"
        .Range(.Cells(F_INI, 9), .Cells(ult, 9)).NumberFormat = "dd-mmm-yy"
        .Range(.Cells(F_INI, 10), .Cells(ult, 11)).NumberFormat = "0.000%"
        .Columns(1).ColumnWidth = 9: .Columns(2).ColumnWidth = 13: .Columns(3).ColumnWidth = 32
        .Columns(4).ColumnWidth = 7: .Columns(5).ColumnWidth = 8: .Columns(6).ColumnWidth = 11
        .Columns(7).ColumnWidth = 13: .Columns(8).ColumnWidth = 13: .Columns(9).ColumnWidth = 11
        .Columns(10).ColumnWidth = 12: .Columns(11).ColumnWidth = 12
    End With
End Sub

'=====================================================================
' CAPA 3: RESUMEN (construido desde cero; idempotente)
'=====================================================================
Public Sub CrearResumen()
    Dim ws As Worksheet, cat As Variant, EQ As String, BQ As String, AX As String
    Dim r As Long, c As Long, i As Long, k As Long, fi As Long, key As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(H_RES)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = H_RES
    End If
    cat = Catalogo()
    HojaBDFMS: HojaEventos
    EQ = H_EVT & "!": BQ = H_BDFMS & "!": AX = "$" & RS_AUX & "$"
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
    ws.Activate

    ws.Range("A1").Value = "Fondos Tradicionales - Resumen por día de marca"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 10

    ' --- fecha de reporte y auxiliares ---
    ws.Range("B3").Value = "Fecha de reporte:": ws.Range("B3").Font.Bold = True
    With ws.Range("C3")
        .Value = Date
        .NumberFormat = "dd-mmm-yy": .Interior.Color = RGB(255, 242, 204): .Font.Bold = True
        .Validation.Delete
        .Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & CLng(DateSerial(2015, 1, 1)), Formula2:="=" & CLng(DateSerial(2035, 12, 31))
        .Validation.InputTitle = "Fecha de reporte"
        .Validation.InputMessage = "Cualquier fecha (ActualizarTodo la pone en hoy). Posicion: ultimo FMS <= fecha. Retornos y contribucion: marcas con Dia <= fecha."
    End With

    Dim bF As String: bF = BQ & "$A$" & F_INI & ":$A$" & F_MAX
    ws.Range(RS_AUX & "3").Formula = "=IFERROR(1/(1/MAXIFS(" & bF & "," & bF & ",""<=""&$C$3)),"""")"                       ' FMS a la fecha
    ws.Range(RS_AUX & "4").Formula = "=IFERROR(1/(1/MAXIFS(" & bF & "," & bF & ",""<""&DATE(YEAR($C$3),MONTH($C$3),1))),"""")"  ' FMS antes del mes
    ws.Range(RS_AUX & "5").Formula = "=IFERROR(1/(1/MAXIFS(" & bF & "," & bF & ",""<""&DATE(YEAR($C$3),1,1))),"""")"           ' FMS antes del anio
    ws.Range(RS_AUX & "6").Formula = "=IFERROR(1/(1/MAXIFS(" & EQ & "$E$" & F_INI & ":$E$" & F_MAX & "," & EQ & "$E$" & F_INI & ":$E$" & F_MAX & ",""<=""&$C$3)),"""")" ' ultima marca
    ws.Range(RS_AUX & "3:" & RS_AUX & "6").NumberFormat = ";;;"

    ws.Range("E3").Value = "Posición FMS al:": ws.Range("E3").Font.Bold = True: ws.Range("E3").HorizontalAlignment = xlRight
    ws.Range("F3").Formula = "=IF(" & AX & "3="""",""sin FMS""," & AX & "3)": ws.Range("F3").NumberFormat = "dd-mmm-yy": ws.Range("F3").Font.Bold = True
    ws.Range("H3").Value = "Última marca:": ws.Range("H3").Font.Bold = True: ws.Range("H3").HorizontalAlignment = xlRight
    ws.Range("I3").Formula = "=IF(" & AX & "6="""",""sin marcas""," & AX & "6)": ws.Range("I3").NumberFormat = "dd-mmm-yy": ws.Range("I3").Font.Bold = True

    ' --- cuadros fila 6-9: retorno del bloque | posicion | contribucion ---
    Dim eS As String, eD As String, eMo As String, eP As String, eCF2 As String, eL As String
    eS = EQ & "$B$" & F_INI & ":$B$" & F_MAX: eD = EQ & "$E$" & F_INI & ":$E$" & F_MAX
    eMo = EQ & "$D$" & F_INI & ":$D$" & F_MAX: eP = EQ & "$P$" & F_INI & ":$P$" & F_MAX
    eCF2 = EQ & "$Q$" & F_INI & ":$Q$" & F_MAX: eL = EQ & "$T$" & F_INI & ":$T$" & F_MAX
    Dim mtd0 As String, ytd0 As String
    mtd0 = "DATE(YEAR($C$3),MONTH($C$3),1)": ytd0 = "DATE(YEAR($C$3),1,1)"

    Dim hdrR As Variant: hdrR = Array("Retorno bloque (Día)", "MTD F1", "YTD F1", "MTD F2", "YTD F2")
    Dim hdrP As Variant: hdrP = Array("Posición", "MM F1", "% F1", "MM F2", "% F2")
    Dim hdrC As Variant: hdrC = Array("Contribución (%)", "MTD F1", "YTD F1", "MTD F2", "YTD F2")
    For k = 0 To 4
        ws.Cells(6, 2 + k).Value = hdrR(k): ws.Cells(6, 8 + k).Value = hdrP(k): ws.Cells(6, 14 + k).Value = hdrC(k)
    Next k
    Encabezado ws, 6, 2, 6: Encabezado ws, 6, 8, 12: Encabezado ws, 6, 14, 18
    Dim mon As Variant: mon = Array("PEN", "USD", "Total")
    For k = 0 To 2
        r = 7 + k
        ws.Cells(r, 2).Value = mon(k): ws.Cells(r, 8).Value = mon(k): ws.Cells(r, 14).Value = mon(k)
        ' contribucion (% del fondo AFP)
        ws.Cells(r, 15).Formula = ContribMon("P", CStr(mon(k)), mtd0, EQ)
        ws.Cells(r, 16).Formula = ContribMon("P", CStr(mon(k)), ytd0, EQ)
        ws.Cells(r, 17).Formula = ContribMon("Q", CStr(mon(k)), mtd0, EQ)
        ws.Cells(r, 18).Formula = ContribMon("Q", CStr(mon(k)), ytd0, EQ)
        ' retorno del bloque = contribucion / peso del bloque al FMS anterior al periodo
        ws.Cells(r, 3).Formula = "=IFERROR(O" & r & "/" & PesoBloque("F1", CStr(mon(k)), AX & "4", BQ) & ","""")"
        ws.Cells(r, 4).Formula = "=IFERROR(P" & r & "/" & PesoBloque("F1", CStr(mon(k)), AX & "5", BQ) & ","""")"
        ws.Cells(r, 5).Formula = "=IFERROR(Q" & r & "/" & PesoBloque("F2", CStr(mon(k)), AX & "4", BQ) & ","""")"
        ws.Cells(r, 6).Formula = "=IFERROR(R" & r & "/" & PesoBloque("F2", CStr(mon(k)), AX & "5", BQ) & ","""")"
        ' posicion a la fecha
        ws.Cells(r, 9).Formula = "=IFERROR(" & PosBloque("G", "F1", CStr(mon(k)), AX & "3", BQ) & "/1000000,"""")"
        ws.Cells(r, 10).Formula = "=IFERROR(" & PesoBloque("F1", CStr(mon(k)), AX & "3", BQ) & ","""")"
        ws.Cells(r, 11).Formula = "=IFERROR(" & PosBloque("G", "F2", CStr(mon(k)), AX & "3", BQ) & "/1000000,"""")"
        ws.Cells(r, 12).Formula = "=IFERROR(" & PesoBloque("F2", CStr(mon(k)), AX & "3", BQ) & ","""")"
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 6)).NumberFormat = "0.00%"
        ws.Cells(r, 9).NumberFormat = "#,##0.0": ws.Cells(r, 11).NumberFormat = "#,##0.0"
        ws.Cells(r, 10).NumberFormat = "0.00%": ws.Cells(r, 12).NumberFormat = "0.00%"
        ws.Range(ws.Cells(r, 15), ws.Cells(r, 18)).NumberFormat = "0.000%"
        ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 8).Font.Bold = True: ws.Cells(r, 14).Font.Bold = True
        If k = 2 Then ws.Range(ws.Cells(r, 2), ws.Cells(r, 18)).Font.Bold = True
    Next k

    ' --- tablas por fondo ---
    Dim hdrT As Variant
    hdrT = Array("Instrumento", "Año-3", "Año-2", "Año-1", "YTD", "1M", "3M", "6M", "12M", "Anualiz.", _
                 "Pos F1 (MM)", "Pos F2 (MM)", "% F1", "% F2", "MTD F1 (pb)", "YTD F1 (pb)", "MTD F2 (pb)", "YTD F2 (pb)", "Últ. marca")
    Dim blq As Variant: blq = Array(Array("FONDOS PEN", RS_PEN_TIT, RS_PEN_INI, 1), Array("FONDOS USD", RS_USD_TIT, RS_USD_INI, 7))
    Dim rIni As Long, rHdr As Long, iIni As Long
    For k = 0 To 1
        ws.Cells(blq(k)(1), 2).Value = blq(k)(0): ws.Cells(blq(k)(1), 2).Font.Bold = True
        rHdr = blq(k)(1) + 1: rIni = blq(k)(2): iIni = blq(k)(3)
        For c = 0 To UBound(hdrT)
            ws.Cells(rHdr, 2 + c).Value = hdrT(c)
        Next c
        For c = 1 To 3
            ws.Cells(rHdr, 2 + c).Formula = "=YEAR($C$3)-" & (4 - c)
            ws.Cells(rHdr, 2 + c).NumberFormat = "0": ws.Cells(rHdr, 2 + c).HorizontalAlignment = xlRight
        Next c
        ws.Cells(rHdr, 11).Formula = "=""Anualiz. ""&YEAR($C$3)"
        Encabezado ws, rHdr, 2, 2 + UBound(hdrT)
        For i = 0 To 5
            fi = iIni + i: r = rIni + i
            ws.Cells(r, 2).Value = cat(fi, 1)
            key = "INDEX(" & H_CONFIG & "!$D$" & CF_CAT_INI & ":$D$" & (CF_CAT_INI + N_FONDOS - 1) & "," & fi & ")"   ' codigo SBS del catalogo
            For c = 1 To 3                                                       ' anios cerrados
                ws.Cells(r, 2 + c).Formula = RetDia(key, "DATE(YEAR($C$3)-" & (4 - c) & ",1,1)", "DATE(YEAR($C$3)-" & (4 - c) & ",12,31)", eS, eD, eL)
            Next c
            ws.Cells(r, 6).Formula = RetDia(key, ytd0, "$C$3", eS, eD, eL)
            ws.Cells(r, 7).Formula = RetDia(key, mtd0, "$C$3", eS, eD, eL)
            ws.Cells(r, 8).Formula = RetDia(key, "DATE(YEAR($C$3),MONTH($C$3)-2,1)", "$C$3", eS, eD, eL)
            ws.Cells(r, 9).Formula = RetDia(key, "DATE(YEAR($C$3),MONTH($C$3)-5,1)", "$C$3", eS, eD, eL)
            ws.Cells(r, 10).Formula = RetDia(key, "DATE(YEAR($C$3),MONTH($C$3)-11,1)", "$C$3", eS, eD, eL)
            ws.Cells(r, 11).Formula = "=IF(F" & r & "="""","""",(1+F" & r & ")^(12/MONTH($C$3))-1)"
            ws.Cells(r, 12).Formula = "=IF(" & AX & "3="""","""",IFERROR(" & PosFondo("G", "F1", key, AX & "3", BQ) & "/1000000,""""))"
            ws.Cells(r, 13).Formula = "=IF(" & AX & "3="""","""",IFERROR(" & PosFondo("G", "F2", key, AX & "3", BQ) & "/1000000,""""))"
            ws.Cells(r, 14).Formula = "=IF(" & AX & "3="""","""",IFERROR(" & PosFondo("G", "F1", key, AX & "3", BQ) & "/" & PosFondo("H", "F1", key, AX & "3", BQ) & ",""""))"
            ws.Cells(r, 15).Formula = "=IF(" & AX & "3="""","""",IFERROR(" & PosFondo("G", "F2", key, AX & "3", BQ) & "/" & PosFondo("H", "F2", key, AX & "3", BQ) & ",""""))"
            ws.Cells(r, 16).Formula = ContribSBS("P", key, mtd0, EQ)
            ws.Cells(r, 17).Formula = ContribSBS("P", key, ytd0, EQ)
            ws.Cells(r, 18).Formula = ContribSBS("Q", key, mtd0, EQ)
            ws.Cells(r, 19).Formula = ContribSBS("Q", key, ytd0, EQ)
            ws.Cells(r, 20).Formula = "=IFERROR(1/(1/MAXIFS(" & eD & "," & eS & "," & key & "," & eD & ",""<=""&$C$3)),"""")"
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 11)).NumberFormat = "0.00%"
            ws.Cells(r, 12).NumberFormat = "#,##0.0": ws.Cells(r, 13).NumberFormat = "#,##0.0"
            ws.Cells(r, 14).NumberFormat = "0.00%": ws.Cells(r, 15).NumberFormat = "0.00%"
            ws.Range(ws.Cells(r, 16), ws.Cells(r, 19)).NumberFormat = "#,##0.0"
            ws.Cells(r, 20).NumberFormat = "dd-mmm-yy"
        Next i
        r = rIni + 6                                                             ' fila total del bloque
        ws.Cells(r, 2).Value = "Total " & Mid$(blq(k)(0), 8)
        For c = 12 To 19
            ws.Cells(r, c).Formula = "=IF(COUNT(" & ColLetra(c) & rIni & ":" & ColLetra(c) & (rIni + 5) & ")=0,"""",SUM(" & _
                                     ColLetra(c) & rIni & ":" & ColLetra(c) & (rIni + 5) & "))"
            ws.Cells(r, c).NumberFormat = IIf(c = 12 Or c = 13, "#,##0.0", IIf(c <= 15, "0.00%", "#,##0.0"))
        Next c
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 20)).Font.Bold = True
        ' heatmaps
        Heatmap ws.Range(ws.Cells(rIni, 3), ws.Cells(rIni + 5, 10))
        For c = 16 To 19: Heatmap ws.Range(ws.Cells(rIni, c), ws.Cells(rIni + 5, c)): Next c
    Next k

    ' --- formato ---
    ws.Columns(1).ColumnWidth = 2: ws.Columns(2).ColumnWidth = 34
    For c = 3 To 20: ws.Columns(c).ColumnWidth = 10: Next c
    ws.Columns(14).ColumnWidth = 15
    ws.Range("B3:B9").Font.Bold = True
    Application.Calculate
    Application.ScreenUpdating = True
    MsgBox "Resumen construido. Cambia la fecha en C3.", vbInformation, "CrearResumen"
End Sub

'--- constructores de formulas del Resumen -----------------------------
Private Function RetDia(key As String, desde As String, hasta As String, eS As String, eD As String, eL As String) As String
    ' retorno del fondo encadenando marcas con Dia en [desde, hasta]; "" si no hubo marcas
    RetDia = "=IF(COUNTIFS(" & eS & "," & key & "," & eD & ","">=""&" & desde & "," & eD & ",""<=""&" & hasta & ")=0,""""," & _
             "EXP(SUMIFS(" & eL & "," & eS & "," & key & "," & eD & ","">=""&" & desde & "," & eD & ",""<=""&" & hasta & "))-1)"
End Function

Private Function ContribSBS(colLetra As String, key As String, desde As String, EQ As String) As String
    ContribSBS = "=SUMIFS(" & EQ & "$" & colLetra & "$" & F_INI & ":$" & colLetra & "$" & F_MAX & "," & _
                 EQ & "$B$" & F_INI & ":$B$" & F_MAX & "," & key & "," & _
                 EQ & "$E$" & F_INI & ":$E$" & F_MAX & ","">=""&" & desde & "," & _
                 EQ & "$E$" & F_INI & ":$E$" & F_MAX & ",""<=""&$C$3)*10000"
End Function

Private Function ContribMon(colLetra As String, moneda As String, desde As String, EQ As String) As String
    Dim crit As String
    If moneda <> "Total" Then crit = "," & EQ & "$D$" & F_INI & ":$D$" & F_MAX & ",""" & moneda & """"
    ContribMon = "=SUMIFS(" & EQ & "$" & colLetra & "$" & F_INI & ":$" & colLetra & "$" & F_MAX & "," & _
                 EQ & "$E$" & F_INI & ":$E$" & F_MAX & ","">=""&" & desde & "," & _
                 EQ & "$E$" & F_INI & ":$E$" & F_MAX & ",""<=""&$C$3" & crit & ")"
End Function

Private Function PosBloque(colLetra As String, fondo As String, moneda As String, celFecha As String, BQ As String) As String
    Dim crit As String
    If moneda <> "Total" Then crit = "," & BQ & "$E$" & F_INI & ":$E$" & F_MAX & ",""" & moneda & """"
    PosBloque = "SUMIFS(" & BQ & "$" & colLetra & "$" & F_INI & ":$" & colLetra & "$" & F_MAX & "," & _
                BQ & "$A$" & F_INI & ":$A$" & F_MAX & "," & celFecha & "," & _
                BQ & "$B$" & F_INI & ":$B$" & F_MAX & ",""" & fondo & """" & crit & ")"
End Function

Private Function PesoBloque(fondo As String, moneda As String, celFecha As String, BQ As String) As String
    PesoBloque = "(" & PosBloque("G", fondo, moneda, celFecha, BQ) & "/AVERAGEIFS(" & BQ & "$H$" & F_INI & ":$H$" & F_MAX & "," & _
                 BQ & "$A$" & F_INI & ":$A$" & F_MAX & "," & celFecha & "," & BQ & "$B$" & F_INI & ":$B$" & F_MAX & ",""" & fondo & """))"
End Function

Private Function PosFondo(colLetra As String, fondo As String, key As String, celFecha As String, BQ As String) As String
    PosFondo = "SUMIFS(" & BQ & "$" & colLetra & "$" & F_INI & ":$" & colLetra & "$" & F_MAX & "," & _
               BQ & "$A$" & F_INI & ":$A$" & F_MAX & "," & celFecha & "," & _
               BQ & "$B$" & F_INI & ":$B$" & F_MAX & ",""" & fondo & """," & _
               BQ & "$C$" & F_INI & ":$C$" & F_MAX & "," & key & ")"
End Function

Private Sub Heatmap(rg As Range)
    rg.FormatConditions.Delete
    With rg.FormatConditions.AddColorScale(3)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 50
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
    End With
End Sub

'=====================================================================
' UTILITARIOS (compartidos)
'=====================================================================
Private Function HojaDatos(nombre As String, titulo As String, hdr As Variant) As Worksheet
    Dim ws As Worksheet, k As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
        ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
        ws.Activate
        ws.Cells(F_INI, 1).Select
        ActiveWindow.FreezePanes = True
    End If
    ws.Range("A1").Value = titulo: ws.Range("A1").Font.Bold = True
    For k = 0 To UBound(hdr)
        ws.Cells(F_HDR, 1 + k).Value = hdr(k)
    Next k
    Encabezado ws, F_HDR, 1, UBound(hdr) + 1
    Set HojaDatos = ws
End Function


Private Function UltimaFila(ws As Worksheet, col As Long) As Long
    UltimaFila = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If UltimaFila < F_INI Then UltimaFila = F_INI - 1
End Function


Private Sub LimpiarDatos(ws As Worksheet, ncol As Long)
    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < F_INI Then ult = F_INI
    ws.Range(ws.Cells(F_INI, 1), ws.Cells(ult + 1, ncol)).Clear
End Sub


Private Sub Encabezado(ws As Worksheet, fila As Long, c1 As Long, c2 As Long)
    With ws.Range(ws.Cells(fila, c1), ws.Cells(fila, c2))
        .Interior.Color = RGB(212, 12, 12)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
End Sub


Private Function ColLetra(c As Long) As String
    ColLetra = Split(ThisWorkbook.Worksheets(1).Cells(1, c).Address(True, False), "$")(0)
End Function


Private Function ListarFMS(desde As Date, dExist As Object, ByRef fechas() As String, _
                           ByRef rutas() As String) As Long
    Dim fso As Object, dArch As Object, partes() As String, r As Long, ruta As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dArch = CreateObject("Scripting.Dictionary")
    partes = Split(CStr(Cfg().Range(CF_CARPETA).Value), ";")
    For r = LBound(partes) To UBound(partes)
        ruta = Trim$(partes(r))
        If Right$(ruta, 1) = "\" Then ruta = Left$(ruta, Len(ruta) - 1)
        If ruta <> vbNullString Then
            If fso.FolderExists(ruta) Then EscanearCarpeta fso.GetFolder(ruta), desde, dExist, dArch
        End If
    Next r
    Dim n As Long, i As Long, j As Long, k As Variant, tmp As String
    n = dArch.Count
    If n = 0 Then Exit Function
    ReDim fechas(1 To n): ReDim rutas(1 To n)
    For Each k In dArch.Keys
        i = i + 1: fechas(i) = CStr(k)
    Next k
    For i = 1 To n - 1
        For j = i + 1 To n
            If fechas(j) < fechas(i) Then tmp = fechas(i): fechas(i) = fechas(j): fechas(j) = tmp
        Next j
    Next i
    For i = 1 To n: rutas(i) = dArch(fechas(i)): Next i
    ListarFMS = n
End Function


Private Sub EscanearCarpeta(fld As Object, desde As Date, dExist As Object, dArch As Object)
    Dim f As Object, sf As Object, nom As String, f8 As String, fecha As Date
    For Each f In fld.Files
        nom = f.Name
        If UCase$(nom) Like "FMS*" And UCase$(nom) Like "*.XLS*" Then
            f8 = ExtraerFecha8(nom)
            If f8 <> vbNullString Then
                fecha = DateSerial(CLng(Left$(f8, 4)), CLng(Mid$(f8, 5, 2)), CLng(Right$(f8, 2)))
                If fecha >= desde And Not dExist.Exists(CLng(fecha)) Then
                    If Not dArch.Exists(f8) Then dArch(f8) = f.Path
                End If
            End If
        End If
    Next f
    On Error Resume Next
    For Each sf In fld.SubFolders
        EscanearCarpeta sf, desde, dExist, dArch
    Next sf
    On Error GoTo 0
End Sub


Private Sub FormatoBDMarcas(ws As Worksheet)
    Dim ult As Long
    ult = UltimaFila(ws, BM_DIA)
    If ult < F_INI Then Exit Sub
    With ws
        .Range(.Cells(F_INI, BM_DIA), .Cells(ult, BM_EEFF)).NumberFormat = "dd-mmm-yy"
        .Range(.Cells(F_INI, BM_VP), .Cells(ult, BM_VP)).NumberFormat = "#,##0.0000"
        .Range(.Cells(F_INI, BM_PCT), .Cells(ult, BM_VA)).NumberFormat = "0.000%"
        .Range(.Cells(F_INI, BM_CARG), .Cells(ult, BM_CARG)).NumberFormat = "dd-mmm-yy hh:mm"
        .Columns(BM_HOJA).ColumnWidth = 6: .Columns(BM_SBS).ColumnWidth = 13: .Columns(BM_INSTR).ColumnWidth = 32
        .Columns(BM_ORDEN).ColumnWidth = 6: .Columns(BM_DIA).ColumnWidth = 10: .Columns(BM_EEFF).ColumnWidth = 10
        .Columns(BM_VP).ColumnWidth = 12: .Columns(BM_PCT).ColumnWidth = 9: .Columns(BM_VA).ColumnWidth = 9
        .Columns(BM_VALIDA).ColumnWidth = 7: .Columns(BM_MOTIVO).ColumnWidth = 22
        .Columns(BM_ARCH).ColumnWidth = 16: .Columns(BM_CARG).ColumnWidth = 14
    End With
End Sub


Private Sub UsadoReal(ws As Worksheet, ByRef ultFila As Long, ByRef ultCol As Long)
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    ultFila = IIf(c Is Nothing, 1, c.Row)
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    ultCol = IIf(c Is Nothing, 1, c.Column)
End Sub


Private Function BuscarCodigo(wsM As Worksheet, codigo As String, ultCol As Long) As Long
    Dim c As Long, v As Variant
    For c = 1 To ultCol
        v = wsM.Cells(MK_FILA_SBS, c).Value
        If VarType(v) = vbString Then
            If UCase$(Trim$(v)) = UCase$(codigo) Then BuscarCodigo = c: Exit Function
        End If
    Next c
End Function


Private Function LeerBloqueMarcas(wsM As Worksheet, col As Long, i As Long, cat As Variant, _
                                  ultFila As Long, ws As Worksheet, ByRef fila As Long, _
                                  archivo As String) As Long
    Dim cVA As Long, k As Long, hdr As String, r As Long, v As Variant, va As Variant
    Dim s As String, n As Long, ee As Variant, vp As Variant, pct As Variant, motivo As String
    cVA = col + MK_OFF_VA
    For k = 0 To 7
        hdr = UCase$(CStr(wsM.Cells(MK_FILA_HDR, col + k).Value))
        If InStr(hdr, "VAR") > 0 And InStr(hdr, "ADJ") > 0 Then cVA = col + k: Exit For
    Next k
    For r = MK_FILA_INI To ultFila
        v = wsM.Cells(r, col).Value
        If IsEmpty(v) Then GoTo Sig
        If VarType(v) = vbDate Then
            If Year(v) >= 2000 And Year(v) <= 2100 Then
                va = wsM.Cells(r, cVA).Value
                ee = wsM.Cells(r, col + MK_OFF_EEFF).Value
                vp = wsM.Cells(r, col + MK_OFF_VP).Value
                pct = wsM.Cells(r, col + MK_OFF_PCT).Value
                motivo = vbNullString
                If Not EsNumero(va) Then motivo = "sin Var Adj"
                If VarType(ee) <> vbDate Then motivo = motivo & IIf(motivo = "", "", "; ") & "sin EEFF"
                ws.Cells(fila, BM_HOJA).Value = wsM.Name
                ws.Cells(fila, BM_SBS).Value = cat(i, 2)
                ws.Cells(fila, BM_INSTR).Value = cat(i, 1)
                ws.Cells(fila, BM_ORDEN).Value = i
                ws.Cells(fila, BM_DIA).Value = CDate(v)
                If VarType(ee) = vbDate Then ws.Cells(fila, BM_EEFF).Value = ee
                If EsNumero(vp) Then ws.Cells(fila, BM_VP).Value = CDbl(vp)
                If EsNumero(pct) Then ws.Cells(fila, BM_PCT).Value = CDbl(pct)
                If EsNumero(va) Then ws.Cells(fila, BM_VA).Value = CDbl(va)
                ws.Cells(fila, BM_VALIDA).Value = IIf(motivo = "", "SI", "NO")
                ws.Cells(fila, BM_MOTIVO).Value = motivo
                ws.Cells(fila, BM_ARCH).Value = archivo
                ws.Cells(fila, BM_CARG).Value = Now
                fila = fila + 1: n = n + 1
            End If
        ElseIf VarType(v) = vbString Then
            s = UCase$(Trim$(v))
            If s = cat(i, 2) Or s = "F1" Or s = "F2" Or s = "F3" Or Left$(s, 6) = "MONEDA" Then Exit For
        End If
Sig:
    Next r
    LeerBloqueMarcas = n
End Function


Private Sub OrdenarRango(ws As Worksheet, ncol As Long, k1 As Long, k2 As Long, k3 As Long)
    Dim ult As Long
    ult = UltimaFila(ws, k1)
    If ult <= F_INI Then Exit Sub
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(F_INI, k1), ws.Cells(ult, k1)), Order:=xlAscending
        If k2 > 0 Then .SortFields.Add Key:=ws.Range(ws.Cells(F_INI, k2), ws.Cells(ult, k2)), Order:=xlAscending
        If k3 > 0 Then .SortFields.Add Key:=ws.Range(ws.Cells(F_INI, k3), ws.Cells(ult, k3)), Order:=xlAscending
        .SetRange ws.Range(ws.Cells(F_INI, 1), ws.Cells(ult, ncol))
        .Header = xlNo
        .Apply
    End With
End Sub


Private Function EsNumero(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Or VarType(v) = vbBoolean Then Exit Function
    If IsError(v) Then Exit Function
    EsNumero = IsNumeric(v)
End Function


Private Function ExtraerFecha8(nombre As String) As String
    Dim i As Long, corrida As Long, f8 As String, aa As Long, mm As Long, dd As Long
    For i = 1 To Len(nombre)
        If Mid$(nombre, i, 1) Like "#" Then
            corrida = corrida + 1
            If corrida >= 8 Then
                f8 = Mid$(nombre, i - 7, 8)
                aa = CLng(Left$(f8, 4)): mm = CLng(Mid$(f8, 5, 2)): dd = CLng(Right$(f8, 2))
                If aa >= 2015 And aa <= 2035 And mm >= 1 And mm <= 12 And dd >= 1 And dd <= 31 Then
                    ExtraerFecha8 = f8: Exit Function
                End If
                dd = CLng(Left$(f8, 2)): mm = CLng(Mid$(f8, 3, 2)): aa = CLng(Right$(f8, 4))
                If aa >= 2015 And aa <= 2035 And mm >= 1 And mm <= 12 And dd >= 1 And dd <= 31 Then
                    ExtraerFecha8 = Format$(aa, "0000") & Format$(mm, "00") & Format$(dd, "00"): Exit Function
                End If
            End If
        Else
            corrida = 0
        End If
    Next i
End Function


Private Function BuscarHoja(wbF As Workbook, nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set BuscarHoja = wbF.Worksheets(nombre)
    On Error GoTo 0
    If Not BuscarHoja Is Nothing Then Exit Function
    For Each ws In wbF.Worksheets
        If InStr(1, ws.Name, nombre, vbTextCompare) > 0 Then Set BuscarHoja = ws: Exit Function
    Next ws
End Function


Private Sub LeerAUM(wsF As Worksheet, dAUM As Object)
    Dim celda As Range, ultFila As Long, r As Long, cCol As Long, txtD As String, fondo As String, v As Variant
    Set celda = wsF.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If celda Is Nothing Then Exit Sub
    ultFila = celda.Row
    For r = 1 To ultFila
        For cCol = 1 To 8
            txtD = Trim$(CStr(wsF.Cells(r, cCol).Value))
            If InStr(1, txtD, MARCA_AUM, vbTextCompare) > 0 Then
                fondo = NormFondo(wsF.Cells(r, FMS_FONDO).Value)
                v = wsF.Cells(r, wsF.Columns.Count).End(xlToLeft).Value
                If fondo <> vbNullString And IsNumeric(v) Then
                    If Not dAUM.Exists(fondo) Then dAUM(fondo) = CDbl(v)
                End If
                Exit For
            End If
        Next cCol
    Next r
End Sub


Private Sub DetectarColumnas(wsF As Worksheet, ByRef cFondo As Long, ByRef cSbs As Long, ByRef cVal As Long)
    Dim c As Long, h As String, fF As Boolean, fS As Boolean, fV As Boolean
    cFondo = FMS_FONDO: cSbs = FMS_SBS: cVal = FMS_VAL
    For c = 1 To 40
        h = UCase$(Trim$(CStr(wsF.Cells(1, c).Value)))
        If h <> vbNullString Then
            If Not fF And InStr(h, "FONDO") > 0 Then cFondo = c: fF = True
            If Not fS And InStr(h, "SBS") > 0 Then cSbs = c: fS = True
            If Not fV And InStr(h, "VAL") > 0 And InStr(h, "TOT") > 0 Then cVal = c: fV = True
        End If
    Next c
End Sub


Private Sub LeerFMS(wsF As Worksheet, dPos As Object, dAUM As Object, ByRef diag As String)
    Dim cFondo As Long, cSbs As Long, cVal As Long
    DetectarColumnas wsF, cFondo, cSbs, cVal
    diag = "hoja " & wsF.Name & ", cols " & ColLetra(cFondo) & "/" & ColLetra(cSbs) & "/" & ColLetra(cVal)
    Dim celda As Range, ultFila As Long, r As Long
    Set celda = wsF.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If celda Is Nothing Then Exit Sub
    ultFila = celda.Row
    Dim txtD As String, txtAlt As String, fondo As String, clave As String, v As Variant
    For r = 2 To ultFila
        txtD = Trim$(CStr(wsF.Cells(r, cSbs).Value))
        txtAlt = vbNullString
        If cSbs <> FMS_SBS Then txtAlt = Trim$(CStr(wsF.Cells(r, FMS_SBS).Value))
        If InStr(1, txtD, MARCA_AUM, vbTextCompare) > 0 Or _
           (txtAlt <> vbNullString And InStr(1, txtAlt, MARCA_AUM, vbTextCompare) > 0) Then
            fondo = NormFondo(wsF.Cells(r, cFondo).Value)
            v = wsF.Cells(r, wsF.Columns.Count).End(xlToLeft).Value
            If fondo <> vbNullString And IsNumeric(v) Then dAUM(fondo) = CDbl(v)
        ElseIf txtD <> vbNullString Then
            v = wsF.Cells(r, cVal).Value
            If IsNumeric(v) And Not IsEmpty(v) Then
                fondo = NormFondo(wsF.Cells(r, cFondo).Value)
                If fondo = "01" Or fondo = "02" Then
                    clave = fondo & "|" & UCase$(txtD)
                    dPos(clave) = dPos(clave) + CDbl(v)
                End If
            End If
        End If
    Next r
End Sub


Private Function NormFondo(v As Variant) As String
    On Error Resume Next
    If Len(Trim$(CStr(v))) > 0 Then
        If IsNumeric(v) Then NormFondo = Format$(CLng(v), "00")
    End If
End Function

