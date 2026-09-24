Attribute VB_Name = "mod_Iliquidos"
'===============================================================================
' mod_Iliquidos  -  v1.0  -  modulo completo
'
' Cuadro "Retornos: Cuasi Sob & Iliquidos".
' Total return por instrumento a partir del vector de renta fija de la SBS,
' con pesos de posicionamiento sacados del FMS.
'
' FUENTES
'   Vector SBS   "AAAAMMDD RFL.xls"   un archivo por dia habil, carpeta Config C2
'   FMS          "FMS_AAAAMMDD.xlsx"  foto de la cartera,        carpeta Config C3
'   Las dos se abren en SOLO LECTURA y se cierran. Nunca se escribe en ellas.
'
' METODOLOGIA FIJA
'   - Precio LIMPIO en % (col J del vector). El cupon se calcula aparte.
'   - Retorno price  = (P_fin - P_base) / P_base
'   - Retorno cupon  = tasa cupon x dias / 365 / P_base
'   - Total return   = la suma de los dos
'   - Current yield  = tasa cupon / P_base
'   - Peso = Val_total del instrumento / total del fondo (Anexo I, codigo 1.1.1)
'   - El codigo del vector trae guiones; se le quitan para cruzar con el FMS.
'   - Nada se corrige en silencio: todo lo raro cae en Mapa como REVISAR.
'
' NADA SE VUELVE A LEER SI YA ESTA
'   - La carpeta del vector se indexa una sola vez por corrida (un Dir, no ~700).
'   - Solo se abren los vectores de fechas que no esten ya en BD_Precios.
'   - El FMS no se reabre si BD_Pesos ya trae esa misma fecha.
'   - BD_Precios se va guardando cada 20 archivos: si se corta la carga, lo leido
'     queda y la siguiente corrida arranca donde quedo.
'   - Recalcular no toca la red: solo rearma Retornos y CUADRO, son segundos.
'
' HOJAS
'   Config . Filtro . Instrumentos . Mapa . BD_Pesos . BD_Precios . Fondos
'   Retornos . CUADRO . Log
'
' MACROS
'   CrearConfig      crea o repara Config e Instrumentos. Conserva lo que ya pusiste.
'   CargarFMS        lee el ultimo FMS <= corte y arma pesos F1 / F2.
'                    Si BD_Pesos ya tiene ese FMS, no abre el archivo.
'   CargarFMSForzado lo relee igual (por si corrigieron el archivo).
'   CargarVectores   lee los vectores que falten y los agrega a BD_Precios.
'   Recalcular       arma Retornos y CUADRO. NO va a la red: son segundos.
'   ActualizarTodo   las tres anteriores, en orden.
'   ArchivarAhora    guarda copia de las BD.
'   CrearFiltro      crea la hoja que decide que se baja del FMS.
'   LimpiarInstrumentos  saca de Instrumentos lo que ya no pasa el filtro.
'                    Nunca borra una fila donde escribiste Nombre o Categoria.
'   ArmarNombres     arma el nombre final: base + cupon + vencimiento.
'   CrearFondos      crea la hoja donde pegas los fondos tradicionales.
'   RefrescarIndice  olvida la lista de archivos de la carpeta del vector.
'   LimpiarNombres   borra nombres definidos heredados de otro libro.
'   GuardarLog       vuelca el ultimo Log a un txt.
'
' ORDEN LA PRIMERA VEZ
'   1. CrearConfig
'   2. Poner las dos carpetas en Config C2 y C3
'   2b. En la hoja Filtro: los codigos SBS que sigues, o los asset class
'   3. CargarFMS            -> llena BD_Pesos e Instrumentos
'   4. En Instrumentos, escribir Nombre y Categoria de los que van al cuadro
'   5. CargarVectores       -> la primera corrida tarda; hace ~230 archivos
'   6. CrearFondos          -> y pegar ahi los fondos tradicionales (opcional)
'   7. Recalcular           -> hojas Retornos y CUADRO
'
' LAS FILAS DE FONDOS TRADICIONALES
'   No salen de aqui: vienen del libro de tradicionales, que ya las calcula desde
'   Marcas. Se pegan (o se enlazan con Pegado especial > Vinculo) en la hoja Fondos
'   y el cuadro las levanta. Asi no hay dos implementaciones de lo mismo que se
'   puedan desincronizar.
'===============================================================================

Option Explicit

Private Const SH_CFG As String = "Config"
Private Const SH_INS As String = "Instrumentos"
Private Const SH_MAP As String = "Mapa"
Private Const SH_PES As String = "BD_Pesos"
Private Const SH_PRE As String = "BD_Precios"
Private Const SH_LOG As String = "Log"
Private Const SH_RET As String = "Retornos"
Private Const SH_CUA As String = "CUADRO"
Private Const SH_FON As String = "Fondos"
Private Const SH_FIL As String = "Filtro"

' BD_Precios
Private Const PR_COLS As Long = 10
' 1 Fecha  2 Codigo SBS  3 Nemonico  4 Moneda  5 Precio limpio %  6 Interes acum
' 7 Tasa cupon  8 Vencimiento  9 TIR %  10 Flag

' BD_Pesos
Private Const PE_COLS As Long = 9
' 1 Codigo SBS  2 Nemonico  3 Emisor  4 Moneda  5 Val_total F1  6 Val_total F2
' 7 Peso F1  8 Peso F2  9 Fecha FMS

Private mT0 As Double
Private mSilencio As Boolean
Private mResumen  As String
Private mFallo    As Boolean
Private mForzar   As Boolean
Private mIdxCarpeta As String
Private mIdxVec     As Object
Private mIdxStamp   As Double
Private mPasos As String
Private mDetalle As String

' CUADRO: filas fijas
'   2 titulo | 4 fecha fin | 5 fecha base | 6 cabecera | 7+ datos
Private Const CU_FIN   As Long = 4
Private Const CU_BASE  As Long = 5
Private Const CU_CAB   As Long = 6
Private Const CU_DAT   As Long = 7

' CUADRO: columnas
'   A oculta moneda | B nombre | C curr yield | D F1 | E F2 | F sep
'   G..L ventanas   | M sep    | N Mensual YTD | O Ann YTD
'   P oculta categoria | Q oculta tipo de fila | R oculta codigo SBS
Private Const CC_MON   As Long = 1
Private Const CC_NOM   As Long = 2
Private Const CC_CY    As Long = 3
Private Const CC_P1    As Long = 4
Private Const CC_SEP   As Long = 6
Private Const CC_W1    As Long = 7
Private Const CC_SEP2  As Long = 13
Private Const CC_MEN   As Long = 14
Private Const CC_ANN   As Long = 15
Private Const CC_CAT   As Long = 16
Private Const CC_TIPO  As Long = 17
Private Const CC_COD   As Long = 18
Private Const N_VENT   As Long = 6
Private Const CAT_FON  As String = "Fondos Tradicionales"


'===============================================================================
'  CONFIG E INSTRUMENTOS
'===============================================================================
Public Sub CrearConfig()
    Dim ws As Worksheet, prev As Object, k As Variant
    Dim scrPrev As Boolean

    scrPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Falla

    Set prev = NuevoDic()
    On Error Resume Next
    For Each k In Array("C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", _
                        "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17")
        prev(CStr(k)) = ThisWorkbook.Worksheets(SH_CFG).Range(CStr(k)).Value
    Next k
    On Error GoTo Falla

    Set ws = HojaOCrea(SH_CFG)
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    ws.Range("B1").Value = "CONFIGURACION - Cuasi Soberanos e Iliquidos"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 13

    Cfg ws, 2, "Carpeta del vector SBS", Rec(prev, "C2", ""), _
        "Donde estan los archivos 'AAAAMMDD RFL.xls'."
    Cfg ws, 3, "Carpeta del FMS", Rec(prev, "C3", ""), _
        "Donde estan los archivos 'FMS_AAAAMMDD.xlsx'."
    Cfg ws, 4, "Fecha de corte", Rec(prev, "C4", Empty), _
        "Vacio = el ultimo vector disponible."
    Cfg ws, 5, "Inicio del historico de precios", Rec(prev, "C5", DateSerial(2025, 10, 31)), _
        "31/10/2025 es la base del FY. Si lo subes, la columna FY sale vacia."
    Cfg ws, 6, "Inicio del periodo libre (MayoTD)", Rec(prev, "C6", DateSerial(2026, 5, 1)), _
        "La base del periodo es esta fecha menos un dia."
    Cfg ws, 7, "Base del FY", Rec(prev, "C7", DateSerial(2025, 10, 31)), "FY = 01 nov a 31 oct."
    Cfg ws, 8, "Dias de la ventana 5D", Rec(prev, "C8", 5), "Dias calendario hacia atras."
    Cfg ws, 9, "Dias sin vector antes de REVISAR", Rec(prev, "C9", 5), _
        "Feriados y fines de semana se saltan solos; mas dias que esto marca REVISAR."
    Cfg ws, 10, "Cargar vectores DESDE", Rec(prev, "C10", Empty), _
        "Vacio = sigue desde la ultima fecha que ya esta en BD_Precios."
    Cfg ws, 11, "Cargar vectores HASTA", Rec(prev, "C11", Empty), _
        "Vacio = la fecha de corte. Sirve para cargar el historico por tramos."
    Cfg ws, 12, "Codigo del Fondo 1", Rec(prev, "C12", "01"), "Columna A de la hoja Cartera."
    Cfg ws, 13, "Codigo del Fondo 2", Rec(prev, "C13", "02"), ""
    Cfg ws, 16, "Total Fondo 1 a mano", Rec(prev, "C16", Empty), _
        "Solo si el Anexo I falla. Vacio = se lee del Anexo I."
    Cfg ws, 17, "Total Fondo 2 a mano", Rec(prev, "C17", Empty), _
        "En soles, el Valor de la Cartera Administrada."
    Cfg ws, 14, "Carpeta de archivo", Rec(prev, "C14", ""), _
        "Vacio = subcarpeta Resumenes junto a este libro."
    Cfg ws, 15, "Formato de archivo", Rec(prev, "C15", "XLSX"), "XLSX / NO"

    ws.Range("B17").Value = "MAPA DEL VECTOR SBS (letras de columna)"
    ws.Range("B17").Font.Bold = True
    Map ws, 18, "Codigo (con guiones)", "A"
    Map ws, 19, "Nemonico", "C"
    Map ws, 20, "Emisor", "E"
    Map ws, 21, "Moneda", "F"
    Map ws, 22, "Precio limpio %", "J"
    Map ws, 23, "Interes acumulado", "M"
    Map ws, 24, "TIR %", "N"
    Map ws, 25, "Fecha vencimiento", "Q"
    Map ws, 26, "Tasa cupon", "R"
    Map ws, 27, "Duracion", "X"

    ws.Range("B29").Value = "MAPA DEL FMS"
    ws.Range("B29").Font.Bold = True
    Map ws, 30, "Codigo fondo", "A"
    Map ws, 31, "Asset class", "C"
    Map ws, 32, "Codigo SBS", "D"
    Map ws, 33, "Emisor", "E"
    Map ws, 34, "Nemonico", "G"
    Map ws, 35, "Moneda", "H"
    Map ws, 36, "Val_total", "J"
    Map ws, 37, "Fecha vencimiento", "L"
    Map ws, 38, "Hoja de la cartera", "Cartera"
    Map ws, 39, "Hoja del anexo", "Anexo I"
    Map ws, 40, "Codigo del total en el anexo", "1.1.1"
    Map ws, 41, "Anexo: columna del fondo", "A"
    Map ws, 42, "Anexo: columna del codigo", "D"
    Map ws, 43, "Anexo: columna del monto", "F"

    ws.Range("C4:C7").NumberFormat = "dd/mm/yyyy"
    ws.Range("C10:C11").NumberFormat = "dd/mm/yyyy"
    ws.Range("C12:C13").NumberFormat = "@"
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 64
    ws.Range("C2:C17").Interior.Color = RGB(255, 242, 204)
    ws.Range("C16:C17").NumberFormat = "#,##0"
    ws.Range("C18:C43").Interior.Color = RGB(255, 242, 204)

    CrearInstrumentos

    Application.ScreenUpdating = scrPrev
    CrearFiltro
    ThisWorkbook.Worksheets(SH_CFG).Activate
    MsgBox "Config, Instrumentos y Filtro listas." & vbCrLf & vbCrLf & _
           "1. Pon las carpetas en C2 (vector) y C3 (FMS)." & vbCrLf & _
           "2. En la hoja Filtro, pega los codigos SBS que sigues" & vbCrLf & _
           "   (o deja los asset class de la columna D)." & vbCrLf & _
           "3. Ejecuta CargarFMS." & vbCrLf & _
           "4. Llena Nombre y Categoria en la hoja Instrumentos." & vbCrLf & _
           "5. Ejecuta CargarVectores.", vbInformation, "CrearConfig"
    Exit Sub

Falla:
    Application.ScreenUpdating = scrPrev
    If Err.Number <> 0 Then MsgBox "CrearConfig se detuvo: " & Err.Description, vbExclamation
End Sub


Public Sub CrearFiltro()
    ' Hoja que decide QUE se baja del FMS. Sin ella se baja toda la cartera.
    ' Si ya existe, no se toca nada de lo que escribiste.
    Dim ws As Worksheet, existia As Boolean, i As Long

    existia = HojaExiste(SH_FIL)
    Set ws = HojaOCrea(SH_FIL)
    If existia Then
        ws.Activate
        Exit Sub
    End If

    ws.Range("A1").Value = "Codigo SBS"
    ws.Range("B1").Value = "Nota (tuya, no se usa)"
    ws.Range("D1").Value = "Asset class que SI se baja"
    ws.Range("E1").Value = "Nota (tuya, no se usa)"
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Range("D2").Value = "Fondo de Inversion Tradicional"
    ws.Range("D3").Value = "Titulos con derecho crediticio"
    ws.Range("D4").Value = "Titulos con derecho de participacion"
    ws.Range("D5").Value = "Bonos de empresas privadas"

    ws.Range("A8").Value = "COMO FUNCIONA"
    ws.Range("A8").Font.Bold = True
    ws.Range("A9").Value = "1. Si la columna A tiene codigos, SOLO esos se bajan. Manda sobre la columna D."
    ws.Range("A10").Value = "2. Si la columna A esta vacia, se baja lo que pertenezca a los asset class de la columna D."
    ws.Range("A11").Value = "3. Si las dos estan vacias, se baja toda la cartera (lo que hacia antes)."
    ws.Range("A12").Value = "4. Un codigo de la columna A que no este en el FMS igual se sigue en el vector,"
    ws.Range("A13").Value = "   y aparece en Instrumentos marcado como 'no esta en el FMS'."
    ws.Range("A14").Value = "5. Las tildes y mayusculas dan igual, y basta con que el asset class del FMS"
    ws.Range("A15").Value = "   contenga el texto que escribas: 'Bonos de empresas' agarra a todos los bonos."
    ws.Range("A17").Value = "La hoja Mapa lista todos los asset class que trae el FMS y cuales pasaron el filtro."
    ws.Range("A17").Value = "Corre CargarFMS y mira ahi para corregir esta lista contra lo que dice el archivo."
    For i = 8 To 17
        ws.Cells(i, 1).Font.Color = RGB(128, 128, 128)
        ws.Cells(i, 1).Font.Italic = True
    Next i

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 40
    ws.Columns("E").ColumnWidth = 30
    ws.Range("A2:A500").NumberFormat = "@"
    CongelarPaneles ws, 1, 0
End Sub


Private Sub LeerFiltro(codes As Object, assets As Object)
    Dim ws As Worksheet, i As Long, ult As Long, t As String
    Set codes = NuevoDic()
    Set assets = NuevoDic()
    If Not HojaExiste(SH_FIL) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_FIL)

    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        t = SinGuiones(Txt(ws.Cells(i, 1).Value))
        If Len(t) > 0 Then codes(t) = 1
    Next i

    ult = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    For i = 2 To ult
        t = NormTxt(Txt(ws.Cells(i, 4).Value))
        If Len(t) > 0 Then assets(t) = Txt(ws.Cells(i, 4).Value)
    Next i
End Sub


Private Function NormMoneda(ByVal t As String) As String
    ' El FMS escribe "S/." y "US$". El cuadro agrupa por este texto, asi que
    ' se normaliza aqui y no en veinte sitios.
    Dim a As String
    a = NormTxt(t)
    a = Replace(Replace(Replace(a, ".", ""), "/", ""), " ", "")
    Select Case a
        Case "S", "SOLES", "PEN", "S1", "NUEVOSOL", "NUEVOSSOLES": NormMoneda = "PEN"
        Case "US$", "US", "USD", "USS", "DOLARES", "DOLAR", "$": NormMoneda = "USD"
        Case ""
            NormMoneda = "SIN MONEDA"
        Case Else
            NormMoneda = UCase$(Trim$(t))
    End Select
End Function


Private Function NormTxt(ByVal t As String) As String
    ' mayusculas, sin tildes, sin espacios de mas: para que "Ilíquidos" e
    ' "ILIQUIDOS" sean lo mismo y no dependa de como lo escriba el FMS
    Dim out As String
    out = UCase$(Trim$(t))
    out = Replace(out, Chr$(193), "A")
    out = Replace(out, Chr$(201), "E")
    out = Replace(out, Chr$(205), "I")
    out = Replace(out, Chr$(211), "O")
    out = Replace(out, Chr$(218), "U")
    out = Replace(out, Chr$(220), "U")
    out = Replace(out, Chr$(209), "N")
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    NormTxt = out
End Function


Private Function PasaFiltro(cod As String, asset As String, _
                            codes As Object, assets As Object) As Boolean
    Dim k As Variant, a As String
    If codes.Count > 0 Then
        PasaFiltro = codes.Exists(cod)
        Exit Function
    End If
    If assets.Count = 0 Then
        PasaFiltro = True
        Exit Function
    End If
    a = NormTxt(asset)
    If Len(a) = 0 Then Exit Function
    For Each k In assets.Keys
        If InStr(1, a, CStr(k), vbBinaryCompare) > 0 Then
            PasaFiltro = True
            Exit Function
        End If
    Next k
End Function


Private Function CensoAssets(v As Variant, cMap As Object, f1 As String, f2 As String, _
                             codes As Object, assets As Object) As Object
    ' asset normalizado -> Array(texto original, filas, pasa Si/No)
    Dim d As Object, i As Long, fo As String, a As String, k As String, arr As Variant
    Set d = NuevoDic()
    Set CensoAssets = d
    If Not IsArray(v) Then Exit Function
    For i = 1 To UBound(v, 1)
        fo = Txt(v(i, cMap("fondo")))
        If MismoFondo(fo, f1) Or MismoFondo(fo, f2) Then
            a = Txt(v(i, cMap("asset")))
            k = NormTxt(a)
            If Len(k) = 0 Then k = "(vacio)"
            If d.Exists(k) Then
                arr = d(k)
                arr(1) = CLng(arr(1)) + 1
                d(k) = arr
            Else
                d(k) = Array(a, 1&, PasaFiltro("", a, codes, assets))
            End If
        End If
    Next i
End Function


Public Sub CrearInstrumentos()
    Dim ws As Worksheet

    Set ws = HojaOCrea(SH_INS)
    If Len(Txt(ws.Range("A1").Value)) > 0 Then Exit Sub   ' ya existe, no se toca

    ws.Range("A1").Value = "Codigo SBS"
    ws.Range("B1").Value = "Nombre base (tuyo)"
    ws.Range("C1").Value = "Categoria"
    ws.Range("D1").Value = "Fuera del Total"
    ws.Range("E1").Value = "Nemonico"
    ws.Range("F1").Value = "Emisor (vector/FMS)"
    ws.Range("G1").Value = "Tasa cupon"
    ws.Range("H1").Value = "Vencimiento"
    ws.Range("I1").Value = "Moneda"
    ws.Range("J1").Value = "Nombre final (lo usa el cuadro)"
    ws.Range("K1").Value = "En el vector"
    ws.Range("A1:K1").Font.Bold = True
    ws.Range("A1:K1").Interior.Color = RGB(217, 217, 217)
    ws.Range("B1:D1").Interior.Color = RGB(255, 242, 204)

    ws.Range("M1").Value = "COLUMNAS QUE LLENAS TU (las demas se refrescan solas)"
    ws.Range("M2").Value = "B  Nombre: manda sobre el automatico. Vacio = se usa el de la columna J."
    ws.Range("M3").Value = "   El vector a veces pone al estructurador y no al riesgo real, por eso existe esta columna."
    ws.Range("M4").Value = "C  Categoria: Cuasi Soberanos / Iliquidos / Coinversiones / Fondos Tradicionales."
    ws.Range("M5").Value = "   Sin categoria, el instrumento no entra al cuadro."
    ws.Range("M6").Value = "D  Fuera del Total: escribe SI para que salga del bloque y aparezca abajo."
    ws.Range("M7").Value = "Esta hoja es append-only: los instrumentos nuevos se agregan al final."
    ws.Range("M1").Font.Bold = True
    ws.Range("M2:M7").Font.Color = RGB(128, 128, 128)

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 28
    ws.Columns("G:I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 34
    ws.Columns("K").ColumnWidth = 12
    ws.Columns("M").ColumnWidth = 96
    CongelarPaneles ws, 1, 1
End Sub


'===============================================================================
'  CARGAR EL FMS  ->  BD_Pesos  +  altas en Instrumentos
'===============================================================================
Public Sub CargarFMS()
    Dim wsCfg As Worksheet, wb As Workbook, wsCar As Worksheet, wsAnx As Worksheet
    Dim carpeta As String, ruta As String, hojaCar As String, hojaAnx As String
    Dim f1 As String, f2 As String, codTot As String
    Dim dCorte As Double, dFms As Double
    Dim v As Variant, va As Variant
    Dim cMap As Object, aMap As Object
    Dim datos As Object, i As Long
    Dim filCod As Object, filAsset As Object, censo As Object
    Dim nSueltos As Long
    Dim diagAnx As Object, wsAnxOk As Boolean, aMano As Boolean
    Dim totF1 As Double, totF2 As Double
    Dim scrPrev As Boolean, calcPrev As XlCalculation, evPrev As Boolean

    mT0 = Timer: mPasos = "": mDetalle = ""
    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Paso "Validando Config"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    carpeta = SinBarra(Txt(wsCfg.Range("C3").Value))
    If Len(carpeta) = 0 Then Err.Raise vbObjectError + 1, , "Falta la carpeta del FMS en Config C3."
    hojaCar = Txt(wsCfg.Range("C38").Value): If Len(hojaCar) = 0 Then hojaCar = "Cartera"
    hojaAnx = Txt(wsCfg.Range("C39").Value): If Len(hojaAnx) = 0 Then hojaAnx = "Anexo I"
    codTot = Txt(wsCfg.Range("C40").Value): If Len(codTot) = 0 Then codTot = "1.1.1"
    f1 = NormFondo(Txt(wsCfg.Range("C12").Value))
    f2 = NormFondo(Txt(wsCfg.Range("C13").Value))

    dCorte = CorteEfectivo(wsCfg)

    Paso "Buscando el ultimo FMS <= corte"
    dFms = BuscarFms(carpeta, dCorte, ruta)
    If dFms = 0 Then
        Err.Raise vbObjectError + 2, , "No encontre ningun FMS en los 20 dias previos a " & _
                  Format$(CDate(dCorte), "dd/mm/yyyy") & vbCrLf & "Carpeta: " & carpeta
    End If

    If Not mForzar Then
        If FmsYaCargado(dFms) Then
            Restaurar calcPrev, scrPrev, evPrev
            Aviso "BD_Pesos ya tiene el FMS del " & Format$(CDate(dFms), "dd/mm/yyyy") & _
                  ", que es el ultimo <= corte." & vbCrLf & _
                  "No se volvio a abrir el archivo." & vbCrLf & vbCrLf & _
                  "Si quieres releerlo igual, corre CargarFMSForzado.", "CargarFMS"
            Exit Sub
        End If
    End If

    Paso "Abriendo el FMS " & Format$(CDate(dFms), "dd/mm/yyyy") & " en solo lectura"
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    On Error Resume Next
    Set wsCar = wb.Worksheets(hojaCar)
    Set wsAnx = wb.Worksheets(hojaAnx)
    On Error GoTo Falla
    If wsCar Is Nothing Then
        wb.Close SaveChanges:=False
        Err.Raise vbObjectError + 3, , "No encontre la hoja '" & hojaCar & "' en el FMS."
    End If

    Paso "Leyendo la cartera a memoria"
    Set cMap = MapaFms(wsCfg)
    v = LeerHoja(wsCar, MaxCol(cMap))
    If Not wsAnx Is Nothing Then
        Set aMap = NuevoDic()
        aMap("fondo") = Letra(Txt(wsCfg.Range("C41").Value), 1)
        aMap("codigo") = Letra(Txt(wsCfg.Range("C42").Value), 4)
        aMap("monto") = Letra(Txt(wsCfg.Range("C43").Value), 6)
        va = LeerHoja(wsAnx, MaxCol(aMap))
    End If

    wsAnxOk = Not (wsAnx Is Nothing)
    wb.Close SaveChanges:=False
    Set wb = Nothing: Set wsCar = Nothing: Set wsAnx = Nothing

    Paso "Aplicando el filtro de la hoja Filtro"
    LeerFiltro filCod, filAsset
    Set censo = CensoAssets(v, cMap, f1, f2, filCod, filAsset)

    Paso "Sumando Val_total por instrumento"
    Set datos = NuevoDic()
    AcumularCartera v, cMap, f1, 1, datos, filCod, filAsset
    AcumularCartera v, cMap, f2, 2, datos, filCod, filAsset

    Paso "Leyendo el total de cada fondo del Anexo I"
    Set diagAnx = RadiografiaAnexo(va, aMap, codTot, hojaAnx, wsAnxOk)
    diagAnx("busco") = "F1=[" & f1 & "]  F2=[" & f2 & "]  (Config C12 y C13)"
    totF1 = Num(wsCfg.Range("C16").Value)
    totF2 = Num(wsCfg.Range("C17").Value)
    If totF1 > 0 Or totF2 > 0 Then aMano = True
    If totF1 <= 0 Then totF1 = TotalAnexo(va, aMap, f1, codTot)
    If totF2 <= 0 Then totF2 = TotalAnexo(va, aMap, f2, codTot)
    If totF1 <= 0 Then totF1 = SumaDic(datos, 1)
    If totF2 <= 0 Then totF2 = SumaDic(datos, 2)

    Paso "Escribiendo BD_Pesos"
    EscribirPesos datos, totF1, totF2, dFms

    Paso "Actualizando la hoja Instrumentos"
    AltasInstrumentos datos
    nSueltos = AltasSueltas(filCod, datos)

    Paso "Escribiendo el Mapa"
    MapaFmsHoja cMap, ruta, dFms, datos.Count, totF1, totF2, censo, filCod, filAsset, _
                nSueltos, diagAnx, aMano

    Restaurar calcPrev, scrPrev, evPrev
    Aviso "v1.0 - CargarFMS terminado en " & Format$(Timer - mT0, "0.0") & " s." & vbCrLf & vbCrLf & _
           mPasos & vbCrLf & _
           "FMS usado: " & Format$(CDate(dFms), "dd/mm/yyyy") & vbCrLf & _
           "Instrumentos que pasaron el filtro: " & datos.Count & vbCrLf & _
           IIf(nSueltos > 0, "Codigos pedidos que NO estan en el FMS: " & nSueltos & vbCrLf, "") & _
           "Filtro: " & DescribeFiltro(filCod, filAsset) & vbCrLf & _
           "Total Fondo 1: " & Format$(totF1 / 1000000, "#,##0.0") & " mn" & vbCrLf & _
           "Total Fondo 2: " & Format$(totF2 / 1000000, "#,##0.0") & " mn" & vbCrLf & vbCrLf & _
           "Ahora llena Nombre y Categoria en la hoja Instrumentos.", "CargarFMS"
    Exit Sub

Falla:
    Reventar wb, calcPrev, scrPrev, evPrev, "CargarFMS"
End Sub


Private Sub AcumularCartera(v As Variant, cMap As Object, fondo As String, _
                            cual As Long, datos As Object, _
                            codes As Object, assets As Object)
    Dim i As Long, cod As String, monto As Double
    Dim arr As Variant

    If Not IsArray(v) Then Exit Sub
    For i = 1 To UBound(v, 1)
        If MismoFondo(Txt(v(i, cMap("fondo"))), fondo) Then
            cod = SinGuiones(Txt(v(i, cMap("codsbs"))))
            If Len(cod) > 0 And Not PasaFiltro(cod, Txt(v(i, cMap("asset"))), codes, assets) Then
                cod = ""
            End If
            If Len(cod) > 0 Then
                monto = Num(v(i, cMap("valtotal")))
                If Not datos.Exists(cod) Then
                    datos(cod) = Array(Txt(v(i, cMap("nemonico"))), _
                                       Txt(v(i, cMap("emisor"))), _
                                       Txt(v(i, cMap("moneda"))), 0#, 0#)
                End If
                arr = datos(cod)
                arr(2 + cual) = Num(arr(2 + cual)) + monto
                datos(cod) = arr
            End If
        End If
    Next i
End Sub


Private Function TotalAnexo(va As Variant, aMap As Object, fondo As String, codTot As String) As Double
    Dim i As Long
    If Not IsArray(va) Then Exit Function
    If aMap Is Nothing Then Exit Function
    For i = 1 To UBound(va, 1)
        If MismoFondo(Txt(va(i, aMap("fondo"))), fondo) Then
            If InStr(1, Txt(va(i, aMap("codigo"))), codTot, vbTextCompare) = 1 Then
                TotalAnexo = Num(va(i, aMap("monto")))
                Exit Function
            End If
        End If
    Next i
End Function


Private Sub EscribirPesos(datos As Object, totF1 As Double, totF2 As Double, dFms As Double)
    Dim ws As Worksheet, out() As Variant, k As Variant, n As Long, arr As Variant
    Dim cab As Variant

    Set ws = HojaOCrea(SH_PES)
    ws.Cells.Clear

    cab = Array("Codigo SBS", "Nemonico", "Emisor", "Moneda", "Val_total F1", _
                "Val_total F2", "Peso F1", "Peso F2", "Fecha FMS")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PE_COLS)).Value = cab
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PE_COLS)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PE_COLS)).Interior.Color = RGB(217, 217, 217)

    If datos.Count = 0 Then Exit Sub
    ReDim out(1 To datos.Count, 1 To PE_COLS)
    n = 0
    For Each k In datos.Keys
        n = n + 1
        arr = datos(k)
        out(n, 1) = "'" & CStr(k)
        out(n, 2) = arr(0)
        out(n, 3) = arr(1)
        out(n, 4) = NormMoneda(CStr(arr(2)))
        out(n, 5) = Num(arr(3))
        out(n, 6) = Num(arr(4))
        out(n, 7) = IIf(totF1 > 0, Num(arr(3)) / totF1, 0)
        out(n, 8) = IIf(totF2 > 0, Num(arr(4)) / totF2, 0)
        out(n, 9) = dFms
    Next k

    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, PE_COLS)).Value = out
    ws.Range(ws.Cells(2, 5), ws.Cells(n + 1, 6)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(2, 7), ws.Cells(n + 1, 8)).NumberFormat = "0.000%"
    ws.Range(ws.Cells(2, 9), ws.Cells(n + 1, 9)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PE_COLS)).EntireColumn.AutoFit
    CongelarPaneles ws, 1, 1

    DefinirNombre "peCod", SH_PES, "$A$2:$A$" & (n + 1)
    DefinirNombre "peNem", SH_PES, "$B$2:$B$" & (n + 1)
    DefinirNombre "pePesoF1", SH_PES, "$G$2:$G$" & (n + 1)
    DefinirNombre "pePesoF2", SH_PES, "$H$2:$H$" & (n + 1)
End Sub


Public Sub CargarFMSForzado()
    ' Relee el FMS aunque BD_Pesos ya lo tenga. Usalo si corrigieron el archivo.
    mForzar = True
    CargarFMS
    mForzar = False
End Sub


Private Function FmsYaCargado(dFms As Double) As Boolean
    Dim ws As Worksheet, ult As Long
    If Not HojaExiste(SH_PES) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_PES)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function
    If Not EsNumero(ws.Cells(2, 9).Value) Then Exit Function
    FmsYaCargado = (CLng(ws.Cells(2, 9).Value) = CLng(dFms))
End Function


Private Sub ArmarNombresSilencioso()
    mSilencio = True
    ArmarNombres
    mSilencio = False
End Sub


Public Sub ArmarNombres()
    ' Llena, desde BD_Precios: G tasa cupon, H vencimiento y J nombre final.
    ' El nombre final se arma solo:   <base> <cupon> <dd/mm/aa>
    ' La base es lo que escribiste en B; si B esta vacia, el emisor; si no,
    ' el nemonico.
    ' Solo pega el cupon si hay cupon, y el vencimiento si hay vencimiento.
    ' Un fondo tradicional no tiene ninguno de los dos: su nombre sale tal
    ' cual lo escribiste, sin nada detras.
    Dim ws As Worksheet, wsP As Worksheet
    Dim i As Long, ult As Long, cod As String, base As String
    Dim datos As Object, arr As Variant
    Dim nArm As Long, nSin As Long, nMio As Long

    On Error GoTo Falla
    If Not HojaExiste(SH_INS) Then
        MsgBox "No existe la hoja " & SH_INS & ".", vbExclamation, "ArmarNombres"
        Exit Sub
    End If
    If Not HojaExiste(SH_PRE) Then
        MsgBox "No existe BD_Precios." & vbCrLf & vbCrLf & _
               "El cupon y el vencimiento salen del vector." & vbCrLf & _
               "Corre CargarVectores y despues vuelve aqui.", vbExclamation, "ArmarNombres"
        Exit Sub
    End If

    ' ultimo cupon y vencimiento por codigo
    Set datos = NuevoDic()
    Set wsP = ThisWorkbook.Worksheets(SH_PRE)
    ult = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuiones(Txt(wsP.Cells(i, 2).Value))
        If Len(cod) > 0 Then
            If Not datos.Exists(cod) Then
                datos(cod) = Array(0#, Num(wsP.Cells(i, 7).Value), Num(wsP.Cells(i, 8).Value))
            End If
            arr = datos(cod)
            If Num(wsP.Cells(i, 1).Value) >= Num(arr(0)) Then
                datos(cod) = Array(Num(wsP.Cells(i, 1).Value), _
                                   Num(wsP.Cells(i, 7).Value), _
                                   Num(wsP.Cells(i, 8).Value))
            End If
        End If
    Next i

    Set ws = ThisWorkbook.Worksheets(SH_INS)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
        If Len(cod) > 0 Then
            base = Trim$(Txt(ws.Cells(i, 2).Value))
            If Len(base) = 0 Then base = Trim$(Txt(ws.Cells(i, 6).Value))
            If Len(base) = 0 Then base = Trim$(Txt(ws.Cells(i, 5).Value))
            If Len(base) = 0 Then base = cod

            If datos.Exists(cod) Then
                arr = datos(cod)
                ws.Cells(i, 7).Value = Num(arr(1))
                ws.Cells(i, 8).Value = Num(arr(2))
                ws.Cells(i, 10).Value = ArmaNombre(base, Num(arr(1)), Num(arr(2)))
                ws.Cells(i, 11).Value = "SI"
                If Num(arr(1)) > 0 Or Num(arr(2)) >= 10000 Then
                    nArm = nArm + 1
                Else
                    nMio = nMio + 1
                End If
            Else
                ws.Cells(i, 10).Value = base
                ws.Cells(i, 11).Value = "REVISAR: no aparece en el vector"
                nSin = nSin + 1
            End If
        End If
    Next i

    ws.Range(ws.Cells(2, 7), ws.Cells(ult, 7)).NumberFormat = "0.000"
    ws.Range(ws.Cells(2, 8), ws.Cells(ult, 8)).NumberFormat = "dd/mm/yyyy"
    ws.Columns(10).ColumnWidth = 36

    If mSilencio Then Exit Sub
    MsgBox "Nombres armados." & vbCrLf & vbCrLf & _
           "Bonos, con cupon y vencimiento del vector: " & nArm & vbCrLf & _
           IIf(nMio > 0, "Sin cupon ni vencimiento (fondos): el nombre queda tal cual: " & _
               nMio & vbCrLf, "") & _
           IIf(nSin > 0, "REVISAR, no estan en el vector: " & nSin & vbCrLf, "") & vbCrLf & _
           "Mira la columna J: eso es lo que va a salir en el cuadro." & vbCrLf & _
           "Si un nombre no te gusta, corrige la columna B y vuelve a correr esto.", _
           vbInformation, "ArmarNombres"
    Exit Sub

Falla:
    MsgBox "ArmarNombres se detuvo: " & Err.Description, vbCritical, "ArmarNombres"
End Sub


Private Function ArmaNombre(base As String, cupon As Double, venc As Double) As String
    ' Solo pega lo que existe. Un fondo no tiene cupon ni vencimiento, asi que
    ' su nombre sale tal cual lo escribiste, sin nada detras.
    Dim out As String
    out = Trim$(base)
    If cupon > 0 Then out = out & " " & Format$(cupon, "0.00")
    If venc >= 10000 Then out = out & " " & Format$(CDate(venc), "dd/mm/yy")
    ArmaNombre = out
End Function


Public Sub Diagnostico()
    ' Un solo dialogo corto que dice por que el filtro no esta haciendo efecto.
    Dim ws As Worksheet, wsCfg As Worksheet
    Dim filCod As Object, filAsset As Object
    Dim nIns As Long, nPes As Long, colA As Long, i As Long, n As Long
    Dim m As String, muestra As String, cAsset As Long

    On Error GoTo Falla
    m = "DIAGNOSTICO" & vbCrLf & String$(34, "-") & vbCrLf

    If HojaExiste(SH_FIL) Then
        LeerFiltro filCod, filAsset
        m = m & "Hoja Filtro: SI existe" & vbCrLf
        m = m & "  codigos en la columna A: " & filCod.Count & vbCrLf
        m = m & "  asset class en la col D: " & filAsset.Count & vbCrLf
    Else
        m = m & "Hoja Filtro: NO EXISTE -> corre CrearFiltro" & vbCrLf
        Set filCod = NuevoDic(): Set filAsset = NuevoDic()
    End If
    m = m & "  filtro efectivo: " & DescribeFiltro(filCod, filAsset) & vbCrLf & vbCrLf

    If HojaExiste(SH_PES) Then
        Set ws = ThisWorkbook.Worksheets(SH_PES)
        nPes = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
    End If
    If HojaExiste(SH_INS) Then
        Set ws = ThisWorkbook.Worksheets(SH_INS)
        nIns = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
    End If
    m = m & "BD_Pesos:     " & nPes & " instrumentos  (lo que paso el filtro)" & vbCrLf
    m = m & "Instrumentos: " & nIns & " filas         (se acumulan, no se borran)" & vbCrLf & vbCrLf

    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    cAsset = Letra(Txt(wsCfg.Range("C31").Value), 3)
    m = m & "Asset class se lee de la columna " & NumALetra(cAsset) & _
            " del FMS (Config C31)." & vbCrLf

    If nPes > 0 Then
        m = m & vbCrLf & "Primeros codigos de BD_Pesos:" & vbCrLf
        Set ws = ThisWorkbook.Worksheets(SH_PES)
        For i = 2 To 6
            If Len(Txt(ws.Cells(i, 1).Value)) > 0 Then
                muestra = muestra & "  " & Txt(ws.Cells(i, 1).Value) & "  " & _
                          Left$(Txt(ws.Cells(i, 3).Value), 26) & vbCrLf
            End If
        Next i
        m = m & muestra
    End If

    m = m & vbCrLf & String$(34, "-") & vbCrLf
    If filCod.Count = 0 And filAsset.Count = 0 Then
        m = m & "QUE HACER: la hoja Filtro esta vacia, por eso baja todo." & vbCrLf & _
                "Llena la columna A o la D y corre CargarFMSForzado."
    ElseIf nPes > 0 And nIns > nPes Then
        m = m & "QUE HACER: el filtro SI funciono (BD_Pesos tiene " & nPes & ")." & vbCrLf & _
                "La lista larga es Instrumentos, que acumula. Corre LimpiarInstrumentos."
    Else
        m = m & "QUE HACER: el filtro no esta dejando fuera nada." & vbCrLf & _
                "Mira la hoja Mapa, la tabla de asset class, y corrige la hoja Filtro."
    End If

    MsgBox m, vbInformation, "Diagnostico"
    Exit Sub

Falla:
    MsgBox "Diagnostico se detuvo: " & Err.Description, vbCritical, "Diagnostico"
End Sub


Public Sub LimpiarInstrumentos()
    ' Saca de la hoja Instrumentos lo que ya no pasa el filtro.
    ' NUNCA borra una fila donde escribiste algo: si tiene Nombre (B),
    ' Categoria (C) o Fuera del Total (D), se queda y se te avisa.
    Dim ws As Worksheet, wsP As Worksheet
    Dim i As Long, ult As Long, cod As String, k As Variant
    Dim vivos As Object, filCod As Object, filAsset As Object
    Dim nBorradas As Long, nSalvadas As Long, resp As VbMsgBoxResult
    Dim quitar() As Long, nQ As Long

    On Error GoTo Falla
    If Not HojaExiste(SH_INS) Then
        MsgBox "No existe la hoja " & SH_INS & ".", vbExclamation, "LimpiarInstrumentos"
        Exit Sub
    End If
    If Not HojaExiste(SH_PES) Then
        MsgBox "No existe BD_Pesos." & vbCrLf & vbCrLf & _
               "Corre CargarFMSForzado con el filtro puesto y despues vuelve aqui.", _
               vbExclamation, "LimpiarInstrumentos"
        Exit Sub
    End If

    ' lo que SI debe quedar: lo que quedo en BD_Pesos mas lo que pediste en Filtro
    Set vivos = NuevoDic()
    Set wsP = ThisWorkbook.Worksheets(SH_PES)
    ult = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuiones(Txt(wsP.Cells(i, 1).Value))
        If Len(cod) > 0 Then vivos(cod) = 1
    Next i
    LeerFiltro filCod, filAsset
    For Each k In filCod.Keys
        vivos(CStr(k)) = 1
    Next k

    If vivos.Count = 0 Then
        MsgBox "BD_Pesos esta vacia. No voy a borrar nada a ciegas.", _
               vbExclamation, "LimpiarInstrumentos"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(SH_INS)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim quitar(1 To ult + 1)
    nQ = 0
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
        If Len(cod) > 0 Then
            If Not vivos.Exists(cod) Then
                If Len(Txt(ws.Cells(i, 2).Value)) > 0 _
                Or Len(Txt(ws.Cells(i, 3).Value)) > 0 _
                Or Len(Txt(ws.Cells(i, 4).Value)) > 0 Then
                    nSalvadas = nSalvadas + 1
                    ws.Cells(i, 11).Value = "REVISAR: ya no pasa el filtro, pero tiene datos tuyos"
                    ws.Range(ws.Cells(i, 1), ws.Cells(i, 11)).Interior.Color = RGB(255, 235, 156)
                Else
                    nQ = nQ + 1
                    quitar(nQ) = i
                End If
            End If
        End If
    Next i

    If nQ = 0 Then
        MsgBox "No hay nada que sacar." & vbCrLf & vbCrLf & _
               IIf(nSalvadas > 0, nSalvadas & " filas no pasan el filtro pero tienen datos " & _
                   "tuyos en B, C o D: quedan marcadas en amarillo para que decidas.", _
                   "Todo lo que esta en Instrumentos pasa el filtro."), _
               vbInformation, "LimpiarInstrumentos"
        Exit Sub
    End If

    resp = MsgBox("Voy a sacar " & nQ & " filas de Instrumentos." & vbCrLf & vbCrLf & _
                  "Son codigos que ya no pasan el filtro y en los que no escribiste " & _
                  "ni Nombre ni Categoria." & vbCrLf & _
                  IIf(nSalvadas > 0, vbCrLf & "Otras " & nSalvadas & " no pasan el filtro pero " & _
                      "SI tienen datos tuyos: esas no las toco, las dejo en amarillo." & vbCrLf, "") & _
                  vbCrLf & "Continuo?", vbYesNo + vbQuestion, "LimpiarInstrumentos")
    If resp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    For i = nQ To 1 Step -1
        ws.Rows(quitar(i)).Delete
        nBorradas = nBorradas + 1
    Next i
    Application.ScreenUpdating = True

    MsgBox "Listo." & vbCrLf & vbCrLf & _
           "Filas sacadas: " & nBorradas & vbCrLf & _
           IIf(nSalvadas > 0, "Filas en amarillo (tenian datos tuyos): " & nSalvadas & vbCrLf, "") & _
           vbCrLf & "BD_Pesos y BD_Precios no se tocaron.", vbInformation, "LimpiarInstrumentos"
    Exit Sub

Falla:
    Application.ScreenUpdating = True
    MsgBox "LimpiarInstrumentos se detuvo: " & Err.Description, vbCritical, "LimpiarInstrumentos"
End Sub


Private Function AltasSueltas(codes As Object, datos As Object) As Long
    ' Codigos que pediste en la hoja Filtro y que el FMS no trae (todavia no los
    ' tienes, o el FMS los llama distinto). Se meten igual a Instrumentos para
    ' que el vector los siga; van sin peso, marcados.
    Dim ws As Worksheet, k As Variant, i As Long, ult As Long
    Dim existentes As Object, cod As String, r As Long

    If codes Is Nothing Then Exit Function
    If codes.Count = 0 Then Exit Function
    If Not HojaExiste(SH_INS) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_INS)
    Set existentes = NuevoDic()
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 1 Then ult = 1
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
        If Len(cod) > 0 Then existentes(cod) = i
    Next i

    r = ult
    For Each k In codes.Keys
        cod = CStr(k)
        If Not datos.Exists(cod) Then
            AltasSueltas = AltasSueltas + 1
            If Not existentes.Exists(cod) Then
                r = r + 1
                ws.Cells(r, 1).Value = "'" & cod
                ws.Cells(r, 6).Value = "REVISAR: pedido en Filtro, no esta en el FMS"
                ws.Cells(r, 6).Interior.Color = RGB(255, 235, 156)
            End If
        End If
    Next k
End Function


Private Function DescribeFiltro(codes As Object, assets As Object) As String
    If codes Is Nothing Then
        DescribeFiltro = "toda la cartera"
    ElseIf codes.Count > 0 Then
        DescribeFiltro = codes.Count & " codigos de la hoja Filtro"
    ElseIf assets.Count > 0 Then
        DescribeFiltro = assets.Count & " asset class de la hoja Filtro"
    Else
        DescribeFiltro = "toda la cartera (la hoja Filtro esta vacia)"
    End If
End Function


Private Sub AltasInstrumentos(datos As Object)
    Dim ws As Worksheet, existentes As Object, k As Variant, arr As Variant
    Dim i As Long, ult As Long, r As Long, cod As String

    CrearInstrumentos
    Set ws = ThisWorkbook.Worksheets(SH_INS)

    Set existentes = NuevoDic()
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 1 Then ult = 1
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
        If Len(cod) > 0 Then existentes(cod) = i
    Next i

    r = ult
    For Each k In datos.Keys
        cod = CStr(k)
        arr = datos(k)
        If existentes.Exists(cod) Then
            i = CLng(existentes(cod))
        Else
            r = r + 1
            i = r
            ws.Cells(i, 1).Value = "'" & cod
        End If
        ' columnas informativas: se refrescan siempre, las tuyas (B, C, D) nunca
        ws.Cells(i, 5).Value = arr(0)
        ws.Cells(i, 6).Value = arr(1)
        ws.Cells(i, 9).Value = NormMoneda(CStr(arr(2)))
    Next k
End Sub


'===============================================================================
'  CARGAR VECTORES  ->  BD_Precios   (incremental y por tramos)
'===============================================================================
Public Sub CargarVectores()
    Dim wsCfg As Worksheet, wsPre As Worksheet
    Dim carpeta As String, ruta As String
    Dim dDesde As Double, dHasta As Double, dFe As Double
    Dim codigos As Object, vMap As Object
    Dim buf() As Variant, nBuf As Long, capa As Long
    Dim nArch As Long, nSalt As Long, nFilas As Long, nHueco As Long, nYa As Long
    Dim ultFecha As Double, tolDias As Long
    Dim ya As Object
    Dim scrPrev As Boolean, calcPrev As XlCalculation, evPrev As Boolean, alertPrev As Boolean
    Dim t As Double

    mT0 = Timer: mPasos = "": mDetalle = ""
    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    alertPrev = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Paso "Validando Config"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    carpeta = SinBarra(Txt(wsCfg.Range("C2").Value))
    If Len(carpeta) = 0 Then Err.Raise vbObjectError + 10, , "Falta la carpeta del vector en Config C2."
    tolDias = CLng(NumDef(wsCfg.Range("C9").Value, 5))
    Set vMap = MapaVector(wsCfg)

    Paso "Armando la lista de codigos a seguir"
    Set codigos = CodigosSeguidos()
    If codigos.Count = 0 Then
        Err.Raise vbObjectError + 11, , _
            "No hay codigos que seguir." & vbCrLf & vbCrLf & _
            "Ejecuta primero CargarFMS, o escribe los codigos a mano en la hoja Instrumentos."
    End If

    Set wsPre = HojaOCrea(SH_PRE)
    PrepararPrecios wsPre
    Set ya = FechasCargadas(wsPre, ultFecha)

    ' --- rango a cargar ---
    If IsDate(wsCfg.Range("C10").Value) Then
        dDesde = CDbl(CDate(wsCfg.Range("C10").Value))
    ElseIf ultFecha > 0 Then
        dDesde = ultFecha + 1
    Else
        dDesde = CDbl(CDate(wsCfg.Range("C5").Value))
    End If

    If IsDate(wsCfg.Range("C11").Value) Then
        dHasta = CDbl(CDate(wsCfg.Range("C11").Value))
    Else
        dHasta = CorteEfectivo(wsCfg)
    End If

    If dHasta < dDesde Then
        Restaurar calcPrev, scrPrev, evPrev
        Application.DisplayAlerts = alertPrev
        Aviso "No hay nada que cargar: BD_Precios ya llega al " & _
               Format$(CDate(ultFecha), "dd/mm/yyyy") & ".", "CargarVectores"
        Exit Sub
    End If

    Paso "Leyendo vectores del " & Format$(CDate(dDesde), "dd/mm/yyyy") & _
         " al " & Format$(CDate(dHasta), "dd/mm/yyyy")

    capa = 60000
    ReDim buf(1 To capa, 1 To PR_COLS)
    nBuf = 0
    nHueco = 0

    For dFe = dDesde To dHasta
        If ya.Exists(CLng(dFe)) Then
            nYa = nYa + 1
        Else
            ruta = RutaVector(carpeta, dFe)
            If Len(ruta) = 0 Then
                nSalt = nSalt + 1
                nHueco = nHueco + 1
            Else
                mDetalle = "Vector " & Format$(CDate(dFe), "dd/mm/yyyy")
                Application.StatusBar = "Vector " & Format$(CDate(dFe), "dd/mm/yyyy") & _
                    "   (" & nArch + 1 & " archivos, " & nFilas & " filas, " & _
                    Format$(Timer - mT0, "0") & " s)"
                nFilas = nFilas + LeerVector(ruta, dFe, vMap, codigos, buf, nBuf, capa, _
                                             IIf(nHueco > tolDias, "REVISAR: " & nHueco & _
                                                 " dias sin vector antes de esta fecha", ""))
                nArch = nArch + 1
                nHueco = 0
                ' Se vuelca cada 20 archivos: si esto se corta a medio camino,
                ' lo ya leido queda guardado y la proxima corrida sigue de ahi.
                If nBuf > 40000 Or (nArch Mod 20) = 0 Then
                    VolcarPrecios wsPre, buf, nBuf
                    nBuf = 0
                End If
                DoEvents
            End If
        End If
    Next dFe

    Paso "Volcando " & nFilas & " filas a BD_Precios"
    VolcarPrecios wsPre, buf, nBuf

    Restaurar calcPrev, scrPrev, evPrev
    Application.DisplayAlerts = alertPrev
    Application.StatusBar = False

    t = Timer - mT0
    Aviso "v1.0 - CargarVectores terminado en " & Format$(t, "0.0") & " s" & _
           IIf(t > 90, " (" & Format$(t / 60, "0.0") & " min)", "") & "." & vbCrLf & vbCrLf & _
           mPasos & vbCrLf & _
           "Archivos leidos ahora: " & nArch & vbCrLf & _
           "Fechas que ya estaban en BD_Precios: " & nYa & vbCrLf & _
           "Fechas sin archivo (fines de semana y feriados): " & nSalt & vbCrLf & _
           "Filas agregadas: " & Format$(nFilas, "#,##0") & vbCrLf & _
           "Codigos seguidos: " & codigos.Count & vbCrLf & _
           "BD_Precios queda con " & FilasPrecios() & " filas en total." & vbCrLf & vbCrLf & _
           IIf(nArch = 0 And nYa > 0, _
               "No leyo ningun archivo porque esas fechas YA estaban cargadas." & vbCrLf & _
               "Para releerlas, borra esas filas de BD_Precios." & vbCrLf & vbCrLf, "") & _
           IIf(nArch = 0 And nYa = 0 And nSalt > 0, _
               "No encontre ningun archivo en la carpeta para esas fechas." & vbCrLf & _
               "Revisa Config C2 y que el archivo se llame 'AAAAMMDD RFL.xls'." & vbCrLf & vbCrLf, "") & _
           "Revisa la hoja Mapa y la columna Flag de BD_Precios.", "CargarVectores"
    Exit Sub

Falla:
    ' no se tira lo ya leido
    On Error Resume Next
    VolcarPrecios wsPre, buf, nBuf
    On Error GoTo 0
    Application.DisplayAlerts = alertPrev
    Reventar Nothing, calcPrev, scrPrev, evPrev, "CargarVectores"
End Sub


Private Function LeerVector(ruta As String, dFe As Double, vMap As Object, codigos As Object, _
                            ByRef buf() As Variant, ByRef nBuf As Long, ByRef capa As Long, _
                            notaHueco As String) As Long
    Dim wb As Workbook, ws As Worksheet, v As Variant
    Dim i As Long, cod As String, n As Long

    On Error GoTo Salir
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    Set ws = wb.Worksheets(1)
    v = LeerHoja(ws, MaxCol(vMap))
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If Not IsArray(v) Then Exit Function

    For i = 1 To UBound(v, 1)
        cod = SinGuiones(Txt(v(i, vMap("codigo"))))
        If Len(cod) > 0 Then
            If codigos.Exists(cod) Then
                nBuf = nBuf + 1
                If nBuf > capa Then
                    capa = capa * 2
                    ReDim Preserve buf(1 To capa, 1 To PR_COLS)
                End If
                buf(nBuf, 1) = dFe
                buf(nBuf, 2) = "'" & cod
                buf(nBuf, 3) = Txt(v(i, vMap("nemonico")))
                buf(nBuf, 4) = Txt(v(i, vMap("moneda")))
                buf(nBuf, 5) = NumOVacio(v(i, vMap("precio")))
                buf(nBuf, 6) = NumOVacio(v(i, vMap("interes")))
                buf(nBuf, 7) = NumOVacio(v(i, vMap("cupon")))
                buf(nBuf, 8) = FechaSegura(v(i, vMap("vencimiento")))
                buf(nBuf, 9) = NumOVacio(v(i, vMap("tir")))
                buf(nBuf, 10) = notaHueco
                If Not EsNumero(v(i, vMap("precio"))) Then
                    buf(nBuf, 10) = MasFlag(buf(nBuf, 10), "REVISAR: sin precio limpio")
                End If
                n = n + 1
            End If
        End If
    Next i
    LeerVector = n
    Exit Function

Salir:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function


Private Sub PrepararPrecios(ws As Worksheet)
    Dim cab As Variant
    If Len(Txt(ws.Range("A1").Value)) > 0 Then Exit Sub
    cab = Array("Fecha", "Codigo SBS", "Nemonico", "Moneda", "Precio limpio %", _
                "Interes acum", "Tasa cupon", "Vencimiento", "TIR %", "Flag")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PR_COLS)).Value = cab
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PR_COLS)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PR_COLS)).Interior.Color = RGB(217, 217, 217)
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 16
    ws.Columns(10).ColumnWidth = 42
    CongelarPaneles ws, 1, 2
End Sub


Private Function FechasCargadas(ws As Worksheet, ByRef ultFecha As Double) As Object
    Dim d As Object, ult As Long, v As Variant, i As Long
    Set d = NuevoDic()
    ultFecha = 0
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then
        Set FechasCargadas = d
        Exit Function
    End If
    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 1)).Value2
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If EsNumero(v(i, 1)) Then
                d(CLng(v(i, 1))) = 1
                If CDbl(v(i, 1)) > ultFecha Then ultFecha = CDbl(v(i, 1))
            End If
        Next i
    Else
        If EsNumero(v) Then
            d(CLng(v)) = 1
            ultFecha = CDbl(v)
        End If
    End If
    Set FechasCargadas = d
End Function


Private Sub VolcarPrecios(ws As Worksheet, buf() As Variant, nBuf As Long)
    Dim ult As Long, i As Long, j As Long, out() As Variant
    If nBuf = 0 Then Exit Sub

    ReDim out(1 To nBuf, 1 To PR_COLS)
    For i = 1 To nBuf
        For j = 1 To PR_COLS
            out(i, j) = buf(i, j)
        Next j
    Next i

    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 1 Then ult = 1
    ws.Range(ws.Cells(ult + 1, 1), ws.Cells(ult + nBuf, PR_COLS)).Value = out
    ' solo se formatea el tramo recien pegado, no toda la hoja
    ws.Range(ws.Cells(ult + 1, 1), ws.Cells(ult + nBuf, 1)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(ult + 1, 8), ws.Cells(ult + nBuf, 8)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(ult + 1, 5), ws.Cells(ult + nBuf, 7)).NumberFormat = "0.000000"
    ws.Range(ws.Cells(ult + 1, 9), ws.Cells(ult + nBuf, 9)).NumberFormat = "0.0000"

    DefinirNombre "prFecha", SH_PRE, "$A$2:$A$" & (ult + nBuf)
    DefinirNombre "prCod", SH_PRE, "$B$2:$B$" & (ult + nBuf)
    DefinirNombre "prPrecio", SH_PRE, "$E$2:$E$" & (ult + nBuf)
    DefinirNombre "prCupon", SH_PRE, "$G$2:$G$" & (ult + nBuf)
End Sub


Private Function FilasPrecios() As Long
    Dim ws As Worksheet, u As Long
    If Not HojaExiste(SH_PRE) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_PRE)
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If u > 1 Then FilasPrecios = u - 1
End Function


Private Function CodigosSeguidos() As Object
    ' codigos de BD_Pesos mas los que tengas escritos a mano en Instrumentos
    Dim d As Object, ws As Worksheet, i As Long, ult As Long, cod As String
    Set d = NuevoDic()

    If HojaExiste(SH_PES) Then
        Set ws = ThisWorkbook.Worksheets(SH_PES)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To ult
            cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
            If Len(cod) > 0 Then d(cod) = 1
        Next i
    End If

    If HojaExiste(SH_INS) Then
        Set ws = ThisWorkbook.Worksheets(SH_INS)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To ult
            cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
            If Len(cod) > 0 Then d(cod) = 1
        Next i
    End If

    Set CodigosSeguidos = d
End Function


'===============================================================================
'  ACTUALIZAR TODO
'===============================================================================
Public Sub ActualizarTodo()
    ' Encadena los tres pasos, se salta lo que ya esta hecho y muestra un solo
    ' dialogo al final. Si un paso se cae, corta ahi y no sigue con el siguiente.
    Dim t As Double
    t = Timer
    mSilencio = True
    mResumen = ""
    mFallo = False

    CargarFMS
    If mFallo Then GoTo Salir
    CargarVectores
    If mFallo Then GoTo Salir
    Recalcular

Salir:
    mSilencio = False
    If mFallo Then Exit Sub          ' el paso que fallo ya mostro su dialogo
    MsgBox "ActualizarTodo termino en " & Format$(Timer - t, "0.0") & " s." & vbCrLf & vbCrLf & _
           mResumen, vbInformation, "ActualizarTodo"
End Sub


'===============================================================================
'  MAPA
'===============================================================================
Private Function RadiografiaAnexo(va As Variant, aMap As Object, codTot As String, _
                                  hojaAnx As String, hojaOk As Boolean) As Object
    ' Deja por escrito que vio en el Anexo I, para que un total que no aparece
    ' se pueda arreglar mirando la hoja Mapa en vez de adivinar.
    Dim d As Object, i As Long, n As Long, nCod As Long
    Dim fo As String, linea As String, fondos As String, muestras As String
    Dim vistos As Object

    Set d = NuevoDic()
    Set RadiografiaAnexo = d
    d("hoja") = hojaAnx
    d("encontrada") = hojaOk
    d("filas") = 0&
    d("con_codigo") = 0&
    d("fondos") = ""
    d("muestra") = ""

    If Not hojaOk Then Exit Function
    If aMap Is Nothing Then Exit Function
    If Not IsArray(va) Then Exit Function

    Set vistos = NuevoDic()
    n = UBound(va, 1)
    d("filas") = CLng(n)
    For i = 1 To n
        fo = Trim$(Txt(va(i, aMap("fondo"))))
        linea = Trim$(Txt(va(i, aMap("codigo"))))
        If Len(fo) > 0 And Len(fo) <= 4 Then
            If Not vistos.Exists(fo) Then
                vistos(fo) = 1
                If Len(fondos) < 60 Then fondos = fondos & IIf(Len(fondos) > 0, " . ", "") & fo
            End If
        End If
        If InStr(1, linea, codTot, vbTextCompare) = 1 Then
            nCod = nCod + 1
            If Len(muestras) < 150 Then
                muestras = muestras & IIf(Len(muestras) > 0, vbLf, "") & _
                           "fila " & i & ": [" & fo & "] [" & Left$(linea, 34) & "] " & _
                           Format$(Num(va(i, aMap("monto"))) / 1000000, "#,##0.0") & " mn"
            End If
        ElseIf nCod = 0 And Len(linea) > 0 And Len(muestras) < 150 And i > 1 And i < 12 Then
            muestras = muestras & IIf(Len(muestras) > 0, vbLf, "") & _
                       "fila " & i & " NO cuadra: [" & Left$(linea, 40) & "]"
        End If
    Next i
    d("con_codigo") = nCod
    d("fondos") = fondos
    d("muestra") = muestras
End Function


Private Sub MapaFmsHoja(cMap As Object, ruta As String, dFms As Double, _
                        nIns As Long, totF1 As Double, totF2 As Double, _
                        censo As Object, codes As Object, assets As Object, _
                        nSueltos As Long, diagAnx As Object, aMano As Boolean)
    Dim ws As Worksheet, r As Long, k As Variant

    Set ws = HojaOCrea(SH_MAP)
    ws.Cells.Clear
    ws.Range("A1").Value = "MAPA Y CONTROLES"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A2").Value = "Corrida:"
    ws.Range("B2").Value = Now
    ws.Range("B2").NumberFormat = "dd/mm/yyyy hh:mm"

    r = 4
    ws.Cells(r, 1).Value = "FMS"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1: ws.Cells(r, 1).Value = "Archivo": ws.Cells(r, 2).Value = ruta
    r = r + 1: ws.Cells(r, 1).Value = "Fecha del FMS": ws.Cells(r, 2).Value = dFms
    ws.Cells(r, 2).NumberFormat = "dd/mm/yyyy"
    r = r + 1: ws.Cells(r, 1).Value = "Instrumentos que pasaron el filtro": ws.Cells(r, 2).Value = nIns
    r = r + 1: ws.Cells(r, 1).Value = "Filtro aplicado": ws.Cells(r, 2).Value = DescribeFiltro(codes, assets)
    If nSueltos > 0 Then
        r = r + 1
        ws.Cells(r, 1).Value = "REVISAR: codigos pedidos que el FMS no trae"
        ws.Cells(r, 2).Value = nSueltos
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 235, 156)
    End If
    r = r + 1: ws.Cells(r, 1).Value = "Total Fondo 1 (mn)": ws.Cells(r, 2).Value = totF1 / 1000000
    ws.Cells(r, 2).NumberFormat = "#,##0.0"
    r = r + 1: ws.Cells(r, 1).Value = "Total Fondo 2 (mn)": ws.Cells(r, 2).Value = totF2 / 1000000
    ws.Cells(r, 2).NumberFormat = "#,##0.0"
    If aMano Then
        r = r + 1
        ws.Cells(r, 1).Value = "Los totales vienen de Config C16 / C17, escritos a mano"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 242, 204)
    End If
    If totF1 <= 0 Or totF2 <= 0 Then
        r = r + 1
        ws.Cells(r, 1).Value = "REVISAR: no encontre el total en el Anexo I; se uso la suma de la cartera"
        ws.Cells(r, 2).Value = "los pesos suman 100% entre estos instrumentos, NO son % del fondo"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 199, 206)
    End If

    r = r + 2
    ws.Cells(r, 1).Value = "ANEXO I  (de aqui sale el total de cada fondo)"
    ws.Cells(r, 1).Font.Bold = True
    If Not diagAnx Is Nothing Then
        r = r + 1: ws.Cells(r, 1).Value = "Hoja buscada": ws.Cells(r, 2).Value = diagAnx("hoja")
        r = r + 1: ws.Cells(r, 1).Value = "La encontre"
        ws.Cells(r, 2).Value = IIf(CBool(diagAnx("encontrada")), "SI", "NO -> revisa Config C39")
        If Not CBool(diagAnx("encontrada")) Then _
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 199, 206)
        r = r + 1: ws.Cells(r, 1).Value = "Filas leidas": ws.Cells(r, 2).Value = diagAnx("filas")
        r = r + 1: ws.Cells(r, 1).Value = "Filas que empiezan con el codigo del total"
        ws.Cells(r, 2).Value = diagAnx("con_codigo")
        If CLng(diagAnx("con_codigo")) = 0 Then _
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 199, 206)
        r = r + 1: ws.Cells(r, 1).Value = "Codigos de fondo que vi en la columna del anexo"
        ws.Cells(r, 2).Value = diagAnx("fondos")
        r = r + 1: ws.Cells(r, 1).Value = "Codigos de fondo que yo busco"
        ws.Cells(r, 2).Value = diagAnx("busco")
        r = r + 1: ws.Cells(r, 1).Value = "Lo que lei (fondo / codigo / monto)"
        ws.Cells(r, 2).Value = diagAnx("muestra")
        ws.Cells(r, 2).WrapText = True
        ws.Rows(r).RowHeight = 62
    End If
    r = r + 1: ws.Cells(r, 1).Value = "Si esto no sale, escribe los totales a mano en Config C16 y C17."
    r = r + 1: ws.Cells(r, 1).Value = "Config C39 hoja / C40 codigo / C41 col fondo / C42 col codigo / C43 col monto"
    ws.Cells(r, 1).Font.Color = RGB(128, 128, 128)

    r = r + 2
    ws.Cells(r, 1).Value = "COLUMNAS DEL FMS QUE SE USARON"
    ws.Cells(r, 1).Font.Bold = True
    For Each k In cMap.Keys
        r = r + 1
        ws.Cells(r, 1).Value = CStr(k)
        ws.Cells(r, 2).Value = NumALetra(CLng(cMap(k)))
    Next k

    r = r + 2
    ws.Cells(r, 1).Value = "ASSET CLASS QUE TRAE EL FMS (F1 + F2)"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1).Value = "Asset class (como lo escribe el FMS)"
    ws.Cells(r, 2).Value = "Filas"
    ws.Cells(r, 3).Value = "Paso el filtro"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    If Not censo Is Nothing Then
        For Each k In censo.Keys
            r = r + 1
            ws.Cells(r, 1).Value = CStr(censo(k)(0))
            ws.Cells(r, 2).Value = CLng(censo(k)(1))
            If codes.Count > 0 Then
                ws.Cells(r, 3).Value = "-"
                ws.Cells(r, 3).Font.Color = RGB(150, 150, 150)
            ElseIf CBool(censo(k)(2)) Then
                ws.Cells(r, 3).Value = "SI"
                ws.Cells(r, 3).Interior.Color = RGB(198, 239, 206)
            Else
                ws.Cells(r, 3).Value = "no"
                ws.Cells(r, 3).Font.Color = RGB(150, 150, 150)
            End If
        Next k
    End If
    If Not codes Is Nothing Then
        If codes.Count > 0 Then
            r = r + 1
            ws.Cells(r, 1).Value = "(la lista de codigos de la hoja Filtro manda: la columna " & _
                                   "de asset class no se uso)"
            ws.Cells(r, 1).Font.Italic = True
            ws.Cells(r, 1).Font.Color = RGB(128, 128, 128)
        End If
    End If
    ws.Columns("C").ColumnWidth = 14

    r = r + 2
    ws.Cells(r, 1).Value = "CONTROLES DE CIERRE"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1: ws.Cells(r, 1).Value = "1. El total del Anexo I sale del codigo 1.1.1, no de la suma de la cartera."
    r = r + 1: ws.Cells(r, 1).Value = "2. Val_total viene en soles: los totales de arriba ya estan en millones."
    r = r + 1: ws.Cells(r, 1).Value = "3. En BD_Precios, filtra la columna Flag: no debe haber 'sin precio limpio'."
    r = r + 1: ws.Cells(r, 1).Value = "4. Los codigos del vector van sin guiones para cruzar con el FMS."
    r = r + 1: ws.Cells(r, 1).Value = "5. La lista de arriba es lo que dice el FMS. Corrige la hoja Filtro contra ella."

    ws.Columns("A").ColumnWidth = 52
    ws.Columns("B").ColumnWidth = 70
End Sub


'===============================================================================
'  ARCHIVO
'===============================================================================
Public Sub ArchivarAhora()
    Dim carpeta As String, ruta As String, base As String, n As Long
    Dim wbNew As Workbook, fmt As String
    Dim scrPrev As Boolean

    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    fmt = UCase$(Txt(ThisWorkbook.Worksheets(SH_CFG).Range("C15").Value))
    If fmt = "NO" Then
        Application.ScreenUpdating = scrPrev
        MsgBox "Config C15 dice NO. No archive nada.", vbInformation, "ArchivarAhora"
        Exit Sub
    End If

    carpeta = SinBarra(Txt(ThisWorkbook.Worksheets(SH_CFG).Range("C14").Value))
    If Len(carpeta) = 0 Then carpeta = ThisWorkbook.Path & Application.PathSeparator & "Resumenes"
    If Dir(carpeta, vbDirectory) = "" Then MkDir carpeta

    base = carpeta & Application.PathSeparator & "Iliquidos_" & Format$(Date, "yyyymmdd")
    ruta = base
    n = 1
    Do While Dir(ruta & ".xlsx") <> ""
        n = n + 1
        ruta = base & "_v" & n
    Loop

    ThisWorkbook.Worksheets(Array(SH_PES, SH_PRE, SH_INS)).Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs Filename:=ruta & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

    Application.ScreenUpdating = scrPrev
    MsgBox "Archivado en:" & vbCrLf & ruta & ".xlsx", vbInformation, "ArchivarAhora"
    Exit Sub

Falla:
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "ArchivarAhora se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "ArchivarAhora"
End Sub


Public Sub LimpiarNombres()
    Dim nm As Name, i As Long, n As Long
    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        If LCase$(Left$(NombreCorto(nm.Name), 2)) = "pr" _
           Or LCase$(Left$(NombreCorto(nm.Name), 2)) = "pe" _
           Or InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then
            On Error Resume Next
            nm.Delete
            On Error GoTo 0
            n = n + 1
        End If
    Next i
    MsgBox "Nombres borrados: " & n, vbInformation, "LimpiarNombres"
End Sub


'===============================================================================
'  RUTAS Y FECHAS
'===============================================================================
Private Function RutaVector(carpeta As String, d As Double) As String
    Dim idx As Object
    Set idx = IndiceVectores(carpeta)
    If idx.Exists(CLng(d)) Then RutaVector = CStr(idx(CLng(d)))
End Function


Private Function IndiceVectores(carpeta As String) As Object
    ' Lee el nombre de TODOS los vectores de la carpeta en una sola pasada y se
    ' lo guarda. Antes preguntaba por cada dia y por cada extension: ~700 golpes
    ' a la red por corrida. Ahora es uno. El indice se reusa 10 minutos.
    Dim d As Object, nom As String, ocho As String, dt As Double

    If Not mIdxVec Is Nothing Then
        If StrComp(mIdxCarpeta, carpeta, vbTextCompare) = 0 Then
            If Timer >= mIdxStamp And Timer - mIdxStamp < 600 Then
                Set IndiceVectores = mIdxVec
                Exit Function
            End If
        End If
    End If

    Set d = NuevoDic()
    mDetalle = "Indexando la carpeta del vector"
    nom = Dir(carpeta & Application.PathSeparator & "*RFL*.xls*")
    Do While Len(nom) > 0
        ocho = OchoDigitos(nom)
        If Len(ocho) = 8 Then
            dt = 0
            On Error Resume Next
            dt = CDbl(DateSerial(CLng(Left$(ocho, 4)), CLng(Mid$(ocho, 5, 2)), CLng(Right$(ocho, 2))))
            On Error GoTo 0
            If dt > 0 Then
                If Not d.Exists(CLng(dt)) Then
                    d(CLng(dt)) = carpeta & Application.PathSeparator & nom
                End If
            End If
        End If
        nom = Dir
    Loop

    Set mIdxVec = d
    mIdxCarpeta = carpeta
    mIdxStamp = Timer
    Set IndiceVectores = d
End Function


Private Function OchoDigitos(ByVal nom As String) As String
    ' Devuelve los primeros 8 digitos seguidos que encuentre en el nombre.
    Dim i As Long, c As String, run As String
    For i = 1 To Len(nom)
        c = Mid$(nom, i, 1)
        If c >= "0" And c <= "9" Then
            run = run & c
            If Len(run) = 8 Then
                OchoDigitos = run
                Exit Function
            End If
        Else
            run = ""
        End If
    Next i
End Function


Public Sub RefrescarIndice()
    ' Usalo solo si copiaron vectores nuevos a la carpeta mientras el libro
    ' estaba abierto y CargarVectores dice que no hay nada nuevo.
    Set mIdxVec = Nothing
    mIdxCarpeta = ""
    MsgBox "Indice de la carpeta borrado." & vbCrLf & _
           "La proxima corrida vuelve a leer la lista de archivos.", vbInformation, "RefrescarIndice"
End Sub


Private Function BuscarFms(carpeta As String, dCorte As Double, ByRef ruta As String) As Double
    Dim i As Long, p As String, ext As Variant
    For i = 0 To 20
        For Each ext In Array(".xlsx", ".xls", ".xlsb")
            p = carpeta & Application.PathSeparator & "FMS_" & _
                Format$(CDate(dCorte - i), "yyyymmdd") & CStr(ext)
            If Dir(p) <> "" Then
                ruta = p
                BuscarFms = dCorte - i
                Exit Function
            End If
        Next ext
    Next i
End Function


Private Function CorteEfectivo(wsCfg As Worksheet) As Double
    Dim carpeta As String, i As Long, dFe As Double
    If IsDate(wsCfg.Range("C4").Value) Then
        CorteEfectivo = CDbl(CDate(wsCfg.Range("C4").Value))
        Exit Function
    End If
    ' vacio: el ultimo vector disponible, buscando hacia atras desde hoy
    carpeta = SinBarra(Txt(wsCfg.Range("C2").Value))
    For i = 0 To 20
        dFe = CDbl(Date) - i
        If Len(RutaVector(carpeta, dFe)) > 0 Then
            CorteEfectivo = dFe
            Exit Function
        End If
    Next i
    CorteEfectivo = CDbl(Date)
End Function


'===============================================================================
'  MAPAS DE COLUMNAS
'===============================================================================
Private Function MapaVector(wsCfg As Worksheet) As Object
    Dim d As Object
    Set d = NuevoDic()
    d("codigo") = Letra(Txt(wsCfg.Range("C18").Value), 1)
    d("nemonico") = Letra(Txt(wsCfg.Range("C19").Value), 3)
    d("emisor") = Letra(Txt(wsCfg.Range("C20").Value), 5)
    d("moneda") = Letra(Txt(wsCfg.Range("C21").Value), 6)
    d("precio") = Letra(Txt(wsCfg.Range("C22").Value), 10)
    d("interes") = Letra(Txt(wsCfg.Range("C23").Value), 13)
    d("tir") = Letra(Txt(wsCfg.Range("C24").Value), 14)
    d("vencimiento") = Letra(Txt(wsCfg.Range("C25").Value), 17)
    d("cupon") = Letra(Txt(wsCfg.Range("C26").Value), 18)
    d("duracion") = Letra(Txt(wsCfg.Range("C27").Value), 24)
    Set MapaVector = d
End Function


Private Function MapaFms(wsCfg As Worksheet) As Object
    Dim d As Object
    Set d = NuevoDic()
    d("fondo") = Letra(Txt(wsCfg.Range("C30").Value), 1)
    d("asset") = Letra(Txt(wsCfg.Range("C31").Value), 3)
    d("codsbs") = Letra(Txt(wsCfg.Range("C32").Value), 4)
    d("emisor") = Letra(Txt(wsCfg.Range("C33").Value), 5)
    d("nemonico") = Letra(Txt(wsCfg.Range("C34").Value), 7)
    d("moneda") = Letra(Txt(wsCfg.Range("C35").Value), 8)
    d("valtotal") = Letra(Txt(wsCfg.Range("C36").Value), 10)
    d("vencimiento") = Letra(Txt(wsCfg.Range("C37").Value), 12)
    Set MapaFms = d
End Function


Private Function MaxCol(m As Object) As Long
    Dim k As Variant
    If m Is Nothing Then Exit Function
    For Each k In m.Keys
        If CLng(m(k)) > MaxCol Then MaxCol = CLng(m(k))
    Next k
End Function


Private Function LeerHoja(ws As Worksheet, maxC As Long) As Variant
    Dim ult As Long
    If maxC < 1 Then maxC = 1
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ult < 1 Then Exit Function
    LeerHoja = ws.Range(ws.Cells(1, 1), ws.Cells(ult, maxC)).Value2
End Function


'===============================================================================
'  UTILIDADES
'===============================================================================
Private Sub Aviso(texto As String, titulo As String)
    ' En ActualizarTodo junta todo en un solo dialogo al final.
    If mSilencio Then
        If Len(mResumen) > 0 Then mResumen = mResumen & vbCrLf & String$(46, "-") & vbCrLf
        mResumen = mResumen & "[" & titulo & "]" & vbCrLf & texto
    Else
        MsgBox texto, vbInformation, titulo
    End If
End Sub


Private Sub Paso(texto As String)
    mPasos = mPasos & "  - " & texto & " (" & Format$(Timer - mT0, "0.0") & " s)" & vbCrLf
    mDetalle = texto
    Application.StatusBar = "Iliquidos: " & texto & "..."
    DoEvents
End Sub


Private Sub Restaurar(calcPrev As XlCalculation, scrPrev As Boolean, evPrev As Boolean)
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = scrPrev
    Application.EnableEvents = evPrev
    Application.StatusBar = False
End Sub


Private Sub Reventar(wb As Workbook, calcPrev As XlCalculation, scrPrev As Boolean, _
                     evPrev As Boolean, quien As String)
    Dim msg As String, nErr As Long, codigo As String, texto As String
    nErr = Err.Number
    msg = Err.Description
    mFallo = True
    mSilencio = False
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False

    codigo = CodigoError(nErr) & " @ " & mDetalle
    texto = quien & " se detuvo." & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & _
            "Paso a paso:" & vbCrLf & mPasos
    GuardarLog "CODIGO: " & codigo & vbCrLf & vbCrLf & texto
    MsgBox "ESCRIBEME ESTA LINEA:" & vbCrLf & vbCrLf & "   " & codigo & vbCrLf & vbCrLf & _
           String$(40, "-") & vbCrLf & msg & vbCrLf & _
           "(el detalle completo quedo en la hoja Log)", vbCritical, quien
End Sub


Private Function CodigoError(n As Long) As String
    If n < 0 Then
        CodigoError = "E" & Format$(n - vbObjectError, "00")
    Else
        CodigoError = CStr(n)
    End If
End Function


Public Sub GuardarLog(texto As String)
    Dim ws As Worksheet, partes As Variant, i As Long, f As Integer, ruta As String
    On Error Resume Next
    Set ws = HojaOCrea(SH_LOG)
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Range("A1").Value = "Ultimo error - " & Format$(Now, "dd/mm/yyyy hh:mm:ss")
        ws.Range("A1").Font.Bold = True
        partes = Split(texto, vbCrLf)
        For i = LBound(partes) To UBound(partes)
            ws.Cells(i + 3, 1).Value = "'" & partes(i)
        Next i
        ws.Columns("A").ColumnWidth = 110
    End If
    If Len(ThisWorkbook.Path) > 0 Then
        ruta = ThisWorkbook.Path & Application.PathSeparator & "Log_Iliquidos.txt"
        f = FreeFile
        Open ruta For Output As #f
        Print #f, "=== " & Format$(Now, "dd/mm/yyyy hh:mm:ss") & " ==="
        Print #f, texto
        Close #f
    End If
    On Error GoTo 0
End Sub


Private Sub Cfg(ws As Worksheet, fila As Long, etiqueta As String, valor As Variant, nota As String)
    ws.Cells(fila, 2).Value = etiqueta
    If Not IsEmpty(valor) Then ws.Cells(fila, 3).Value = valor
    ws.Cells(fila, 4).Value = nota
    ws.Cells(fila, 4).Font.Color = RGB(128, 128, 128)
End Sub


Private Sub Map(ws As Worksheet, fila As Long, campo As String, valor As String)
    ws.Cells(fila, 2).Value = campo
    ws.Cells(fila, 3).Value = valor
End Sub


Private Function Rec(prev As Object, celda As String, porDefecto As Variant) As Variant
    If prev.Exists(celda) Then
        If Not IsEmpty(prev(celda)) Then
            If Len(Trim$(Txt(prev(celda)))) > 0 Then
                Rec = prev(celda)
                Exit Function
            End If
        End If
    End If
    Rec = porDefecto
End Function


Private Function NuevoDic() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set NuevoDic = d
End Function


Private Function HojaOCrea(nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
    End If
    Set HojaOCrea = ws
End Function


Private Function HojaExiste(nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    HojaExiste = Not ws Is Nothing
End Function


Private Sub CongelarPaneles(ws As Worksheet, filas As Long, cols As Long)
    Dim wsPrev As Object
    On Error Resume Next
    ThisWorkbook.Activate
    Set wsPrev = ActiveSheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = filas
    ActiveWindow.SplitColumn = cols
    ActiveWindow.FreezePanes = True
    If Not wsPrev Is Nothing Then wsPrev.Activate
    On Error GoTo 0
End Sub


Private Sub DefinirNombre(nombre As String, hoja As String, refLocal As String)
    Dim ws As Worksheet, rng As Range, nm As Name, i As Long
    Set ws = ThisWorkbook.Worksheets(hoja)
    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        If StrComp(NombreCorto(nm.Name), nombre, vbTextCompare) = 0 Then
            On Error Resume Next
            nm.Delete
            On Error GoTo 0
        End If
    Next i
    Set rng = ws.Range(refLocal)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=nombre, RefersTo:=rng
    On Error GoTo 0
End Sub


Private Function NombreCorto(s As String) As String
    Dim p As Long
    p = InStrRev(s, "!")
    If p > 0 Then
        NombreCorto = Mid$(s, p + 1)
    Else
        NombreCorto = s
    End If
End Function


Private Function Letra(ByVal s As String, porDefecto As Long) As Long
    Dim i As Long, n As Long, ch As Long
    s = UCase$(Trim$(s))
    If Len(s) = 0 Then
        Letra = porDefecto
        Exit Function
    End If
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1)) - 64
        If ch < 1 Or ch > 26 Then
            Letra = porDefecto
            Exit Function
        End If
        n = n * 26 + ch
    Next i
    Letra = n
End Function


Private Function NumALetra(ByVal n As Long) As String
    Dim s As String, r As Long
    Do While n > 0
        r = ((n - 1) Mod 26)
        s = Chr$(65 + r) & s
        n = (n - r - 1) \ 26
    Loop
    NumALetra = s
End Function


Private Function SinGuiones(s As String) As String
    SinGuiones = Replace(Replace(Trim$(s), "-", ""), " ", "")
End Function


Private Function SinBarra(ByVal s As String) As String
    s = Trim$(s)
    Do While Len(s) > 0 And Right$(s, 1) = Application.PathSeparator
        s = Left$(s, Len(s) - 1)
    Loop
    SinBarra = s
End Function


Private Function MismoFondo(ByVal a As String, ByVal b As String) As Boolean
    ' "01" = "1" = 1 = " 01 " = "F1" ... todo lo que en el fondo es el mismo fondo.
    Dim x As String, y As String
    x = SoloDigitos(a)
    y = SoloDigitos(b)
    If Len(x) > 0 And Len(y) > 0 Then
        MismoFondo = (CDbl(x) = CDbl(y))
    Else
        MismoFondo = (NormFondo(a) = NormFondo(b))
    End If
End Function


Private Function SoloDigitos(ByVal t As String) As String
    Dim i As Long, c As String, out As String
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c >= "0" And c <= "9" Then out = out & c
    Next i
    If Len(out) > 15 Then out = Left$(out, 15)
    SoloDigitos = out
End Function


Private Function NormFondo(ByVal s As String) As String
    ' "09" -> "9",  "01" -> "1"
    s = Trim$(s)
    Do While Len(s) > 1 And Left$(s, 1) = "0"
        s = Mid$(s, 2)
    Loop
    NormFondo = s
End Function


Private Function Txt(v As Variant) As String
    If IsError(v) Then
        Txt = ""
    ElseIf IsEmpty(v) Then
        Txt = ""
    ElseIf IsNull(v) Then
        Txt = ""
    Else
        Txt = CStr(v)
    End If
End Function


Private Function EsNumero(v As Variant) As Boolean
    If IsError(v) Then
        EsNumero = False
    ElseIf IsEmpty(v) Then
        EsNumero = False
    ElseIf VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then
            EsNumero = False
        Else
            EsNumero = IsNumeric(v)
        End If
    Else
        EsNumero = IsNumeric(v)
    End If
End Function


Private Function Num(v As Variant) As Double
    If EsNumero(v) Then Num = CDbl(v)
End Function


Private Function NumDef(v As Variant, porDefecto As Double) As Double
    If EsNumero(v) Then
        NumDef = CDbl(v)
    Else
        NumDef = porDefecto
    End If
End Function


Private Function NumOVacio(v As Variant) As Variant
    If EsNumero(v) Then
        NumOVacio = CDbl(v)
    Else
        NumOVacio = Empty
    End If
End Function


Private Function FechaSegura(v As Variant) As Variant
    ' el FMS trae AAAAMMDD como texto; el vector trae fecha de Excel
    Dim s As String
    If EsNumero(v) Then
        If CDbl(v) > 19000101 And CDbl(v) < 21001231 Then
            s = Format$(CDbl(v), "00000000")
            FechaSegura = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Right$(s, 2)))
        Else
            FechaSegura = CDbl(v)
        End If
        Exit Function
    End If
    s = Trim$(Txt(v))
    If Len(s) = 8 And IsNumeric(s) Then
        FechaSegura = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Right$(s, 2)))
    ElseIf IsDate(s) Then
        FechaSegura = CDate(s)
    Else
        FechaSegura = Empty
    End If
End Function


Private Function MasFlag(actual As Variant, nuevo As String) As String
    Dim a As String
    a = Txt(actual)
    If Len(a) = 0 Then
        MasFlag = nuevo
    Else
        MasFlag = a & " | " & nuevo
    End If
End Function


Private Function SumaDic(datos As Object, cual As Long) As Double
    Dim k As Variant, arr As Variant
    For Each k In datos.Keys
        arr = datos(k)
        SumaDic = SumaDic + Num(arr(2 + cual))
    Next k
End Function


'===============================================================================
'  HOJA Retornos  Y  CUADRO
'===============================================================================


Public Sub Recalcular()
    Dim wsCfg As Worksheet, dCorte As Double
    Dim ins As Object, orden As Variant
    Dim scrPrev As Boolean, calcPrev As XlCalculation, evPrev As Boolean

    mT0 = Timer: mPasos = "": mDetalle = ""
    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Paso "Validando"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    If Not HojaExiste(SH_PRE) Then Err.Raise vbObjectError + 20, , _
        "No hay BD_Precios. Ejecuta CargarVectores."
    dCorte = CorteBD(wsCfg)
    If dCorte <= 0 Then Err.Raise vbObjectError + 21, , _
        "BD_Precios esta vacia. Ejecuta CargarVectores."

    Paso "Armando los nombres"
    ArmarNombresSilencioso

    Paso "Leyendo la hoja Instrumentos"
    Set ins = LeerInstrumentos(orden)
    If ins.Count = 0 Then Err.Raise vbObjectError + 22, , _
        "Ningun instrumento tiene Categoria en la hoja Instrumentos." & vbCrLf & _
        "Sin categoria no entran al cuadro."

    Paso "Armando la hoja Retornos"
    ArmarRetornos ins, orden, dCorte, wsCfg

    Paso "Armando el CUADRO"
    ArmarCuadro ins, orden, dCorte

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    Restaurar calcPrev, scrPrev, evPrev

    Aviso "v1.0 - Recalcular terminado en " & Format$(Timer - mT0, "0.0") & " s." & vbCrLf & vbCrLf & _
           mPasos & vbCrLf & _
           "Instrumentos en el cuadro: " & ins.Count & vbCrLf & _
           "Corte: " & Format$(CDate(dCorte), "dd/mm/yyyy") & vbCrLf & vbCrLf & _
           "Revisa la columna Flag de la hoja Retornos.", "Recalcular"
    Exit Sub

Falla:
    Reventar Nothing, calcPrev, scrPrev, evPrev, "Recalcular"
End Sub


Private Function CorteBD(wsCfg As Worksheet) As Double
    Dim ws As Worksheet, ult As Long
    If IsDate(wsCfg.Range("C4").Value) Then
        CorteBD = CDbl(CDate(wsCfg.Range("C4").Value))
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(SH_PRE)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function
    CorteBD = Application.WorksheetFunction.Max(ws.Range("A2:A" & ult))
End Function


Private Function LeerInstrumentos(ByRef orden As Variant) As Object
    ' Nombre -> Array(codigo, nombre, categoria, moneda, fuera)
    ' orden = los codigos en el orden en que aparecen en la hoja
    Dim ws As Worksheet, d As Object, i As Long, ult As Long
    Dim cod As String, nom As String, cat As String, mon As String
    Dim lst() As String, n As Long

    Set d = NuevoDic()
    orden = Array()
    If Not HojaExiste(SH_INS) Then
        Set LeerInstrumentos = d
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SH_INS)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then
        Set LeerInstrumentos = d
        Exit Function
    End If

    ReDim lst(0 To ult)
    n = -1
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
        cat = Trim$(Txt(ws.Cells(i, 3).Value))
        If Len(cod) > 0 And Len(cat) > 0 Then
            ' manda la columna J (nombre final, que arma ArmarNombres)
            nom = Trim$(Txt(ws.Cells(i, 10).Value))
            If Len(nom) = 0 Then nom = Trim$(Txt(ws.Cells(i, 2).Value))
            If Len(nom) = 0 Then nom = Trim$(Txt(ws.Cells(i, 5).Value))
            If Len(nom) = 0 Then nom = cod
            mon = NormMoneda(Txt(ws.Cells(i, 9).Value))
            If Not d.Exists(cod) Then
                d(cod) = Array(cod, nom, cat, mon, _
                               UCase$(Left$(Trim$(Txt(ws.Cells(i, 4).Value)) & "  ", 2)) = "SI")
                n = n + 1
                lst(n) = cod
            End If
        End If
    Next i

    If n >= 0 Then
        ReDim Preserve lst(0 To n)
        orden = lst
    End If
    Set LeerInstrumentos = d
End Function


'--------------------------------------------------------------------- Fondos
Public Sub CrearFondos()
    ' Hoja que el usuario pega o vincula desde el libro de fondos tradicionales.
    ' El modulo NUNCA la escribe: solo la lee al armar el CUADRO.
    Dim ws As Worksheet, cab As Variant, existia As Boolean

    existia = HojaExiste(SH_FON)
    Set ws = HojaOCrea(SH_FON)
    If existia Then
        MsgBox "La hoja " & SH_FON & " ya existe y no se toca." & vbCrLf & vbCrLf & _
               "Pega ahi los fondos tradicionales desde el otro libro." & vbCrLf & _
               "Recalcular los lee y los mete al cuadro como una categoria mas.", _
               vbInformation, "Fondos"
        ws.Activate
        Exit Sub
    End If

    cab = Array("Nombre", "Moneda", "Peso F1", "Peso F2", "Curr. Yield", _
                "WTD", "5D", "MTD", "MayoTD", "YTD", "FY")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).Value = cab
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 11))
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("A3").Value = "Pega aqui (o vincula con =) los fondos tradicionales:"
    ws.Range("A4").Value = "nombre, moneda PEN/USD, pesos y retornos ya calculados, en tanto por uno."
    ws.Range("A5").Value = "Curr. Yield puede ir vacio. Deja la fila 2 libre o usala, da igual."
    ws.Range(ws.Cells(3, 1), ws.Cells(5, 1)).Font.Color = RGB(128, 128, 128)
    ws.Range(ws.Cells(3, 1), ws.Cells(5, 1)).Font.Italic = True
    ws.Columns(1).ColumnWidth = 38
    ws.Columns(2).ColumnWidth = 9
    ws.Range(ws.Columns(3), ws.Columns(11)).ColumnWidth = 10
    ws.Range(ws.Cells(2, 3), ws.Cells(500, 11)).NumberFormat = "0.00%"
    CongelarPaneles ws, 1, 1
    ws.Activate

    MsgBox "Hoja " & SH_FON & " creada." & vbCrLf & vbCrLf & _
           "Pega ahi los fondos tradicionales (nombre, moneda, pesos y retornos)." & vbCrLf & _
           "Si la dejas vacia, el cuadro simplemente no muestra esa categoria.", _
           vbInformation, "Fondos"
End Sub


Private Function LeerFondos() As Collection
    ' Devuelve una coleccion de Array(nombre, moneda, f1, f2, cy, r0..r5)
    Dim ws As Worksheet, c As New Collection, i As Long, ult As Long, j As Long
    Dim nom As String, mon As String, v(0 To 10) As Variant

    Set LeerFondos = c
    If Not HojaExiste(SH_FON) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_FON)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function

    For i = 2 To ult
        nom = Trim$(Txt(ws.Cells(i, 1).Value))
        mon = UCase$(Trim$(Txt(ws.Cells(i, 2).Value)))
        If Len(nom) > 0 And (mon = "PEN" Or mon = "USD" Or Len(mon) = 3) Then
            v(0) = nom
            v(1) = mon
            For j = 2 To 10
                If IsNumeric(ws.Cells(i, j + 1).Value) Then
                    v(j) = CDbl(ws.Cells(i, j + 1).Value)
                Else
                    v(j) = Empty
                End If
            Next j
            c.Add Array(v(0), v(1), v(2), v(3), v(4), v(5), v(6), v(7), v(8), v(9), v(10))
        End If
    Next i
End Function


Private Sub FilaFondo(ws As Worksheet, r As Long, mon As String, f As Variant)
    Dim j As Long, colYTD As String

    mDetalle = "CUADRO: fondo tradicional " & f(0) & " en la fila " & r
    ws.Cells(r, CC_MON).Value = mon
    ws.Cells(r, CC_CAT).Value = CAT_FON
    ws.Cells(r, CC_NOM).Value = f(0)
    ws.Cells(r, CC_NOM).IndentLevel = 2
    ws.Cells(r, CC_TIPO).Value = "INS"
    If Not IsEmpty(f(4)) Then ws.Cells(r, CC_CY).Value = f(4)
    If Not IsEmpty(f(2)) Then ws.Cells(r, CC_P1).Value = f(2)
    If Not IsEmpty(f(3)) Then ws.Cells(r, CC_P1 + 1).Value = f(3)
    For j = 0 To N_VENT - 1
        If Not IsEmpty(f(5 + j)) Then ws.Cells(r, CC_W1 + j).Value = f(5 + j)
    Next j

    colYTD = NumALetra(CC_W1 + 4)
    ws.Cells(r, CC_ANN).Formula = _
        "=IFERROR((1+" & colYTD & r & ")^(365/(" & colYTD & "$" & CU_FIN & "-" & _
        colYTD & "$" & CU_BASE & "))-1,NA())"
    ws.Cells(r, CC_MEN).Formula = "=IFERROR(" & NumALetra(CC_ANN) & r & "/12,NA())"
End Sub


'-------------------------------------------------------------------- Retornos
Private Sub ArmarRetornos(ins As Object, orden As Variant, dCorte As Double, wsCfg As Worksheet)
    Dim ws As Worksheet, r As Long, i As Long, j As Long
    Dim vent As Variant, bases() As Double
    Dim arr As Variant, cab As Variant
    Dim tol As Long

    vent = Array("WTD", "5D", "MTD", "MayoTD", "YTD", "FY")
    bases = BasesVentana(wsCfg, dCorte)
    tol = CLng(NumDef(wsCfg.Range("C9").Value, 5))

    Set ws = HojaOCrea(SH_RET)
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0
    ws.Cells.Clear

    cab = Array("Codigo SBS", "Nemonico", "Nombre", "Categoria", "Moneda", "Ventana", _
                "Fecha base", "Fecha fin", "Base efectiva", "Fin efectiva", _
                "P base", "P fin", "Dias", "Tasa cupon", _
                "Ret. price", "Ret. cupon", "SUMA", "Curr. yield", "Flag")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 19)).Value = cab
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 19)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 19)).Interior.Color = RGB(64, 64, 64)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 19)).Font.Color = RGB(255, 255, 255)

    r = 1
    If IsArray(orden) Then
        For i = LBound(orden) To UBound(orden)
            arr = ins(orden(i))
            For j = 0 To N_VENT - 1
                r = r + 1
                mDetalle = "Retornos: " & arr(1) & " / " & vent(j)
                ws.Cells(r, 1).Value = "'" & arr(0)
                ws.Cells(r, 2).Formula = "=IFERROR(INDEX(peNem,MATCH($A" & r & ",peCod,0)),"""")"
                ws.Cells(r, 3).Value = arr(1)
                ws.Cells(r, 4).Value = arr(2)
                ws.Cells(r, 5).Value = arr(3)
                ws.Cells(r, 6).Value = vent(j)
                ws.Cells(r, 7).Value = bases(j)
                ws.Cells(r, 8).Value = dCorte
                ws.Cells(r, 9).Formula = _
                    "=IFERROR(MAXIFS(prFecha,prCod,$A" & r & ",prFecha,""<=""&G" & r & "),"""")"
                ws.Cells(r, 10).Formula = _
                    "=IFERROR(MAXIFS(prFecha,prCod,$A" & r & ",prFecha,""<=""&H" & r & "),"""")"
                ws.Cells(r, 11).Formula = _
                    "=IFERROR(AVERAGEIFS(prPrecio,prCod,$A" & r & ",prFecha,I" & r & "),"""")"
                ws.Cells(r, 12).Formula = _
                    "=IFERROR(AVERAGEIFS(prPrecio,prCod,$A" & r & ",prFecha,J" & r & "),"""")"
                ws.Cells(r, 13).Formula = "=IF(OR(I" & r & "="""",J" & r & "=""""),"""",J" & r & "-I" & r & ")"
                ws.Cells(r, 14).Formula = _
                    "=IFERROR(AVERAGEIFS(prCupon,prCod,$A" & r & ",prFecha,J" & r & "),"""")"
                ws.Cells(r, 15).Formula = "=IFERROR((L" & r & "-K" & r & ")/K" & r & ","""")"
                ws.Cells(r, 16).Formula = "=IFERROR(N" & r & "*M" & r & "/365/K" & r & ","""")"
                ws.Cells(r, 17).Formula = "=IFERROR(O" & r & "+P" & r & ","""")"
                ws.Cells(r, 18).Formula = "=IFERROR(N" & r & "/L" & r & ","""")"
                ws.Cells(r, 19).Formula = _
                    "=IF(OR(K" & r & "="""",L" & r & "=""""),""REVISAR: falta precio""," & _
                    "IF(G" & r & "-I" & r & ">" & tol & ",""REVISAR: la base usa un vector ""&(G" & r & _
                    "-I" & r & ")&"" dias anterior""," & _
                    "IF(H" & r & "-J" & r & ">" & tol & ",""REVISAR: el fin usa un vector ""&(H" & r & _
                    "-J" & r & ")&"" dias anterior"","""")))"
            Next j
        Next i
    End If

    ws.Range(ws.Cells(2, 7), ws.Cells(r, 10)).NumberFormat = "dd/mm/yy"
    ws.Range(ws.Cells(2, 11), ws.Cells(r, 12)).NumberFormat = "0.0000"
    ws.Range(ws.Cells(2, 14), ws.Cells(r, 14)).NumberFormat = "0.000"
    ws.Range(ws.Cells(2, 15), ws.Cells(r, 18)).NumberFormat = "0.00%"
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 14
    ws.Columns(3).ColumnWidth = 32
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 9
    ws.Columns(6).ColumnWidth = 9
    ws.Range(ws.Columns(7), ws.Columns(18)).ColumnWidth = 11
    ws.Columns(19).ColumnWidth = 46
    CongelarPaneles ws, 1, 3

    If r > 1 Then
        DefinirNombre "reCod", SH_RET, "$A$2:$A$" & r
        DefinirNombre "reVent", SH_RET, "$F$2:$F$" & r
        DefinirNombre "reSuma", SH_RET, "$Q$2:$Q$" & r
        DefinirNombre "reCY", SH_RET, "$R$2:$R$" & r
    End If
End Sub


Private Function BasesVentana(wsCfg As Worksheet, dCorte As Double) As Double()
    Dim b(0 To 5) As Double, dc As Date
    dc = CDate(dCorte)
    b(0) = dCorte - (Weekday(dc, vbMonday) - 1) - 1              ' WTD: domingo anterior
    b(1) = dCorte - NumDef(wsCfg.Range("C8").Value, 5)           ' 5D
    b(2) = CDbl(DateSerial(Year(dc), Month(dc), 0))              ' MTD
    If IsDate(wsCfg.Range("C6").Value) Then
        b(3) = CDbl(CDate(wsCfg.Range("C6").Value)) - 1          ' MayoTD
    Else
        b(3) = dCorte
    End If
    b(4) = CDbl(DateSerial(Year(dc), 1, 1)) - 1                  ' YTD
    If IsDate(wsCfg.Range("C7").Value) Then
        b(5) = CDbl(CDate(wsCfg.Range("C7").Value))              ' FY
    Else
        b(5) = CDbl(DateSerial(Year(dc) - 1, 10, 31))
    End If
    BasesVentana = b
End Function


'---------------------------------------------------------------------- CUADRO
Private Sub ArmarCuadro(ins As Object, orden As Variant, dCorte As Double)
    Dim ws As Worksheet, r As Long, r1 As Long, ultFila As Long
    Dim grupos As Object, monedas As Variant, cats As Variant
    Dim im As Long, ic As Long, i As Long
    Dim arr As Variant, mon As String, cat As String
    Dim vent As Variant, j As Long, bases() As Double
    Dim fondos As Collection, fo As Variant

    Set fondos = LeerFondos()
    vent = Array("WTD", "5D", "MTD", "MayoTD", "YTD", "FY")
    bases = BasesVentana(ThisWorkbook.Worksheets(SH_CFG), dCorte)

    Set ws = HojaOCrea(SH_CUA)
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    ws.Range("B2").Value = "Retornos: Cuasi Sob & Iliquidos"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 15
    ws.Range("B2").Font.Color = RGB(192, 0, 0)
    ws.Cells(2, CC_W1).Value = dCorte
    ws.Cells(2, CC_W1).NumberFormat = "dd/mm/yyyy"
    ws.Cells(2, CC_W1).Font.Bold = True
    ws.Cells(2, CC_W1 + 1).Value = "<- corte (Config C4, vacio = ultima fecha de BD_Precios)"
    ws.Cells(2, CC_W1 + 1).Font.Color = RGB(128, 128, 128)
    ws.Range("B3").Value = "Total return = variacion de precio limpio + cupon devengado (dias/365)" & _
        "   |   peso = Val_total / total del fondo (FMS)   |   filas grises: promedio ponderado F1+F2"
    ws.Range("B3").Font.Color = RGB(128, 128, 128)
    ws.Range("B3").Font.Size = 8

    ' --- encabezado de ventanas ---
    ws.Cells(CU_FIN, CC_NOM).Value = "fin"
    ws.Cells(CU_BASE, CC_NOM).Value = "base"
    ws.Range(ws.Cells(CU_FIN, CC_NOM), ws.Cells(CU_BASE, CC_NOM)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(CU_FIN, CC_NOM), ws.Cells(CU_BASE, CC_NOM)).Font.Color = RGB(128, 128, 128)
    For j = 0 To N_VENT - 1
        ws.Cells(CU_FIN, CC_W1 + j).Value = dCorte
        ws.Cells(CU_BASE, CC_W1 + j).Value = bases(j)
        ws.Cells(CU_CAB, CC_W1 + j).Value = vent(j)
    Next j
    ws.Cells(CU_CAB, CC_CY).Value = "Curr. Yield"
    ws.Cells(CU_CAB, CC_P1).Value = "F1"
    ws.Cells(CU_CAB, CC_P1 + 1).Value = "F2"
    ws.Cells(CU_CAB, CC_MEN).Value = "Mensual YTD"
    ws.Cells(CU_CAB, CC_ANN).Value = "Ann YTD"

    ' --- filas ---
    Set grupos = NuevoDic()
    monedas = OrdenMonedas(ins, orden, fondos)
    r1 = CU_DAT
    r = CU_DAT - 1

    For im = LBound(monedas) To UBound(monedas)
        mon = CStr(monedas(im))
        r = r + 1
        grupos(r) = "BLQ"
        ws.Cells(r, CC_MON).Value = mon
        ws.Cells(r, CC_NOM).Value = mon
        ws.Cells(r, CC_TIPO).Value = "BLQ"

        cats = OrdenCategorias(ins, orden, mon)
        For ic = LBound(cats) To UBound(cats)
            cat = CStr(cats(ic))
            r = r + 1
            grupos(r) = "SUB"
            ws.Cells(r, CC_MON).Value = mon
            ws.Cells(r, CC_CAT).Value = cat
            ws.Cells(r, CC_NOM).Value = cat
            ws.Cells(r, CC_TIPO).Value = "SUB"

            For i = LBound(orden) To UBound(orden)
                arr = ins(orden(i))
                If arr(3) = mon And arr(2) = cat And Not arr(4) Then
                    r = r + 1
                    FilaInstrumento ws, r, mon, cat, CStr(arr(0)), CStr(arr(1)), "INS"
                End If
            Next i
        Next ic

        ' fondos tradicionales de esa moneda, pegados en la hoja Fondos
        If CuentaFondos(fondos, mon) > 0 Then
            r = r + 1
            grupos(r) = "SUB"
            ws.Cells(r, CC_MON).Value = mon
            ws.Cells(r, CC_CAT).Value = CAT_FON
            ws.Cells(r, CC_NOM).Value = CAT_FON
            ws.Cells(r, CC_TIPO).Value = "SUB"
            For Each fo In fondos
                If CStr(fo(1)) = mon Then
                    r = r + 1
                    FilaFondo ws, r, mon, fo
                End If
            Next fo
        End If
    Next im

    ' --- Total ---
    r = r + 1
    grupos(r) = "TOT"
    ws.Cells(r, CC_NOM).Value = "Total"
    ws.Cells(r, CC_TIPO).Value = "TOT"

    ' --- fuera del Total ---
    Dim hayFuera As Boolean
    For i = LBound(orden) To UBound(orden)
        arr = ins(orden(i))
        If arr(4) Then hayFuera = True
    Next i
    If hayFuera Then
        r = r + 2
        ws.Cells(r, CC_NOM).Value = "Fuera del Total"
        ws.Cells(r, CC_NOM).Font.Italic = True
        ws.Cells(r, CC_NOM).Font.Size = 8
        ws.Cells(r, CC_NOM).Font.Color = RGB(128, 128, 128)
        ws.Cells(r, CC_TIPO).Value = "SEP"
        For i = LBound(orden) To UBound(orden)
            arr = ins(orden(i))
            If arr(4) Then
                r = r + 1
                FilaInstrumento ws, r, "", "", CStr(arr(0)), CStr(arr(1)), "OUT"
            End If
        Next i
    End If

    ultFila = r

    ' --- filas de grupo, ya con el rango completo ---
    For Each arr In grupos.Keys
        FilaGrupo ws, CLng(arr), CStr(grupos(arr)), r1, ultFila
    Next arr

    FormatearCuadro ws, ultFila
End Sub


Private Sub FilaInstrumento(ws As Worksheet, r As Long, mon As String, cat As String, _
                            cod As String, nom As String, tipo As String)
    Dim j As Long, cl As String, colYTD As String

    mDetalle = "CUADRO: " & nom & " en la fila " & r
    ws.Cells(r, CC_MON).Value = mon
    ws.Cells(r, CC_CAT).Value = cat
    ws.Cells(r, CC_NOM).Value = nom
    ws.Cells(r, CC_TIPO).Value = tipo
    ws.Cells(r, CC_COD).Value = "'" & cod
    If tipo = "INS" Then ws.Cells(r, CC_NOM).IndentLevel = 2

    ws.Cells(r, CC_CY).Formula = _
        "=IFERROR(AVERAGEIFS(reCY,reCod,$R" & r & ",reVent,""YTD""),"""")"
    ws.Cells(r, CC_P1).Formula = "=IFERROR(SUMIFS(pePesoF1,peCod,$R" & r & "),0)"
    ws.Cells(r, CC_P1 + 1).Formula = "=IFERROR(SUMIFS(pePesoF2,peCod,$R" & r & "),0)"

    For j = 0 To N_VENT - 1
        cl = NumALetra(CC_W1 + j)
        ws.Cells(r, CC_W1 + j).Formula = _
            "=IFERROR(AVERAGEIFS(reSuma,reCod,$R" & r & ",reVent," & cl & "$" & CU_CAB & "),"""")"
    Next j

    colYTD = NumALetra(CC_W1 + 4)
    ws.Cells(r, CC_ANN).Formula = _
        "=IFERROR((1+" & colYTD & r & ")^(365/(" & colYTD & "$" & CU_FIN & "-" & _
        colYTD & "$" & CU_BASE & "))-1,NA())"
    ws.Cells(r, CC_MEN).Formula = "=IFERROR(" & NumALetra(CC_ANN) & r & "/12,NA())"
End Sub


Private Sub FilaGrupo(ws As Worksheet, r As Long, tipo As String, r1 As Long, r2 As Long)
    Dim j As Long, w As String, cond As String

    mDetalle = "CUADRO: formulas del grupo en la fila " & r
    w = "($" & NumALetra(CC_P1) & "$" & r1 & ":$" & NumALetra(CC_P1) & "$" & r2 & "+$" & _
        NumALetra(CC_P1 + 1) & "$" & r1 & ":$" & NumALetra(CC_P1 + 1) & "$" & r2 & ")"
    cond = CondGrupo(tipo, r, r1, r2)

    ws.Cells(r, CC_P1).Formula = SumPeso(tipo, r, r1, r2, CC_P1)
    ws.Cells(r, CC_P1 + 1).Formula = SumPeso(tipo, r, r1, r2, CC_P1 + 1)
    ws.Cells(r, CC_CY).Formula = "=" & Prom(w, cond, CC_CY, r1, r2)
    For j = 0 To N_VENT - 1
        ws.Cells(r, CC_W1 + j).Formula = "=" & Prom(w, cond, CC_W1 + j, r1, r2)
    Next j
    ws.Cells(r, CC_MEN).Formula = "=" & Prom(w, cond, CC_MEN, r1, r2)
    ws.Cells(r, CC_ANN).Formula = "=" & Prom(w, cond, CC_ANN, r1, r2)
End Sub


Private Function Prom(w As String, cond As String, col As Long, r1 As Long, r2 As Long) As String
    ' Promedio ponderado por F1+F2 sobre las filas INS del grupo.
    ' ISNUMBER deja fuera del numerador Y del denominador a los instrumentos
    ' sin dato en esa ventana, para que un hueco no arrastre el promedio a cero.
    Dim rg As String, base As String
    rg = NumALetra(col) & "$" & r1 & ":" & NumALetra(col) & "$" & r2
    base = w & "*" & cond & "*ISNUMBER(" & rg & ")"
    Prom = "IFERROR(SUMPRODUCT(" & base & "," & rg & ")/SUMPRODUCT(" & base & "),"""")"
End Function


Private Function CondGrupo(tipo As String, r As Long, r1 As Long, r2 As Long) As String
    Dim soloIns As String, rgT As String
    rgT = "$" & NumALetra(CC_TIPO) & "$" & r1 & ":$" & NumALetra(CC_TIPO) & "$" & r2
    soloIns = "(" & rgT & "=""INS"")"
    Select Case tipo
        Case "BLQ"
            CondGrupo = "($" & NumALetra(CC_MON) & "$" & r1 & ":$" & NumALetra(CC_MON) & "$" & r2 & _
                        "=$" & NumALetra(CC_MON) & r & ")*" & soloIns
        Case "SUB"
            CondGrupo = "($" & NumALetra(CC_MON) & "$" & r1 & ":$" & NumALetra(CC_MON) & "$" & r2 & _
                        "=$" & NumALetra(CC_MON) & r & ")*($" & NumALetra(CC_CAT) & "$" & r1 & ":$" & _
                        NumALetra(CC_CAT) & "$" & r2 & "=$" & NumALetra(CC_CAT) & r & ")*" & soloIns
        Case Else
            CondGrupo = soloIns
    End Select
End Function


Private Function SumPeso(tipo As String, r As Long, r1 As Long, r2 As Long, col As Long) As String
    Dim rgP As String, rgT As String, rgM As String, rgC As String
    rgP = NumALetra(col) & "$" & r1 & ":" & NumALetra(col) & "$" & r2
    rgT = "$" & NumALetra(CC_TIPO) & "$" & r1 & ":$" & NumALetra(CC_TIPO) & "$" & r2
    rgM = "$" & NumALetra(CC_MON) & "$" & r1 & ":$" & NumALetra(CC_MON) & "$" & r2
    rgC = "$" & NumALetra(CC_CAT) & "$" & r1 & ":$" & NumALetra(CC_CAT) & "$" & r2
    Select Case tipo
        Case "BLQ"
            SumPeso = "=SUMIFS(" & rgP & "," & rgM & ",$" & NumALetra(CC_MON) & r & "," & rgT & ",""INS"")"
        Case "SUB"
            SumPeso = "=SUMIFS(" & rgP & "," & rgM & ",$" & NumALetra(CC_MON) & r & "," & rgC & _
                      ",$" & NumALetra(CC_CAT) & r & "," & rgT & ",""INS"")"
        Case Else
            SumPeso = "=SUMIFS(" & rgP & "," & rgT & ",""INS"")"
    End Select
End Function


Private Function OrdenMonedas(ins As Object, orden As Variant, fondos As Collection) As Variant
    Dim d As Object, i As Long, arr As Variant, k As Variant
    Dim res() As String, n As Long
    Set d = NuevoDic()
    If IsArray(orden) Then
        If UBound(orden) >= LBound(orden) Then
            For i = LBound(orden) To UBound(orden)
                arr = ins(orden(i))
                If Not arr(4) Then d(CStr(arr(3))) = 1
            Next i
        End If
    End If
    If Not fondos Is Nothing Then
        For Each arr In fondos
            d(CStr(arr(1))) = 1
        Next arr
    End If
    If d.Count = 0 Then
        OrdenMonedas = Array()
        Exit Function
    End If
    ReDim res(0 To d.Count - 1)
    n = -1
    If d.Exists("PEN") Then n = n + 1: res(n) = "PEN"
    If d.Exists("USD") Then n = n + 1: res(n) = "USD"
    For Each k In d.Keys
        If CStr(k) <> "PEN" And CStr(k) <> "USD" Then n = n + 1: res(n) = CStr(k)
    Next k
    OrdenMonedas = res
End Function


Private Function CuentaFondos(fondos As Collection, mon As String) As Long
    Dim fo As Variant, n As Long
    If fondos Is Nothing Then Exit Function
    For Each fo In fondos
        If CStr(fo(1)) = mon Then n = n + 1
    Next fo
    CuentaFondos = n
End Function


Private Function OrdenCategorias(ins As Object, orden As Variant, mon As String) As Variant
    Dim d As Object, i As Long, arr As Variant, k As Variant
    Dim res() As String, n As Long
    Set d = NuevoDic()
    For i = LBound(orden) To UBound(orden)
        arr = ins(orden(i))
        If Not arr(4) And CStr(arr(3)) = mon Then d(CStr(arr(2))) = 1
    Next i
    If d.Count = 0 Then
        OrdenCategorias = Array()
        Exit Function
    End If
    ReDim res(0 To d.Count - 1)
    n = -1
    For Each k In d.Keys
        n = n + 1
        res(n) = CStr(k)
    Next k
    OrdenCategorias = res
End Function


'--------------------------------------------------------------------- formato
Private Sub FormatearCuadro(ws As Worksheet, ultFila As Long)
    Dim r As Long, j As Long, tipo As String, rFila As Range

    If ultFila < CU_DAT Then Exit Sub

    ws.Columns(CC_MON).ColumnWidth = 12
    ws.Columns(CC_NOM).ColumnWidth = 38
    ws.Columns(CC_CY).ColumnWidth = 11
    ws.Range(ws.Columns(CC_P1), ws.Columns(CC_P1 + 1)).ColumnWidth = 8
    ws.Columns(CC_SEP).ColumnWidth = 1.6
    ws.Range(ws.Columns(CC_W1), ws.Columns(CC_W1 + N_VENT - 1)).ColumnWidth = 10
    ws.Columns(CC_SEP2).ColumnWidth = 1.6
    ws.Range(ws.Columns(CC_MEN), ws.Columns(CC_ANN)).ColumnWidth = 12

    With ws.Range(ws.Cells(CU_FIN, CC_W1), ws.Cells(CU_BASE, CC_W1 + N_VENT - 1))
        .NumberFormat = "dd/mm/yy"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(150, 150, 150)
    End With
    ws.Range(ws.Cells(CU_DAT, CC_CY), ws.Cells(ultFila, CC_P1 + 1)).NumberFormat = "0.00%;-0.00%;"
    ws.Range(ws.Cells(CU_DAT, CC_W1), ws.Cells(ultFila, CC_W1 + N_VENT - 1)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(CU_DAT, CC_MEN), ws.Cells(ultFila, CC_ANN)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(CU_DAT, CC_CY), ws.Cells(ultFila, CC_ANN)).HorizontalAlignment = xlCenter

    With ws.Range(ws.Cells(CU_CAB, CC_W1), ws.Cells(CU_CAB, CC_W1 + N_VENT - 1))
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(CU_CAB, CC_MEN), ws.Cells(CU_CAB, CC_ANN))
        .Interior.Color = RGB(127, 127, 127)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(CU_CAB, CC_CY), ws.Cells(CU_CAB, CC_P1 + 1))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
    End With

    With ws.Range(ws.Cells(CU_DAT, CC_NOM), ws.Cells(ultFila, CC_ANN)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With

    For r = CU_DAT To ultFila
        tipo = Txt(ws.Cells(r, CC_TIPO).Value)
        Set rFila = ws.Range(ws.Cells(r, CC_NOM), ws.Cells(r, CC_ANN))
        Select Case tipo
            Case "BLQ"
                rFila.Interior.Color = RGB(128, 128, 128)
                rFila.Font.Color = RGB(255, 255, 255)
                rFila.Font.Bold = True
            Case "SUB"
                rFila.Interior.Color = RGB(217, 217, 217)
                rFila.Font.Bold = True
                ws.Cells(r, CC_NOM).IndentLevel = 1
            Case "TOT"
                rFila.Font.Bold = True
                rFila.Borders(xlEdgeTop).LineStyle = xlContinuous
                rFila.Borders(xlEdgeTop).Weight = xlMedium
                rFila.Borders(xlEdgeBottom).LineStyle = xlContinuous
                rFila.Borders(xlEdgeBottom).Weight = xlMedium
            Case "SEP", ""
                rFila.Borders.LineStyle = xlNone
        End Select
    Next r

    For j = 0 To N_VENT - 1
        Escala RangoIns(ws, CC_W1 + j, CU_DAT, ultFila)
    Next j
    Escala RangoIns(ws, CC_MEN, CU_DAT, ultFila)
    Escala RangoIns(ws, CC_ANN, CU_DAT, ultFila)

    ws.Columns(CC_MON).Hidden = True
    ws.Columns(CC_CAT).Hidden = True
    ws.Columns(CC_TIPO).Hidden = True
    ws.Columns(CC_COD).Hidden = True
    CongelarPaneles ws, CU_CAB, CC_NOM
End Sub


Private Function RangoIns(ws As Worksheet, col As Long, r1 As Long, r2 As Long) As Range
    Dim r As Long, u As Range, t As String
    For r = r1 To r2
        t = Txt(ws.Cells(r, CC_TIPO).Value)
        If t = "INS" Or t = "OUT" Then
            If u Is Nothing Then
                Set u = ws.Cells(r, col)
            Else
                Set u = Application.Union(u, ws.Cells(r, col))
            End If
        End If
    Next r
    Set RangoIns = u
End Function


Private Sub Escala(rng As Range)
    ' divergente y simetrica en cero: el 0.00% queda blanco, no rojo
    Dim fc As ColorScale, c As Range
    Dim vals() As Double, n As Long, i As Long, k As Long, t As Double, M As Double

    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    rng.FormatConditions.Delete
    On Error GoTo 0

    ReDim vals(1 To rng.Cells.Count)
    n = 0
    For Each c In rng.Cells
        If Not IsError(c.Value) Then
            If IsNumeric(c.Value) And Len(Txt(c.Text)) > 0 Then
                n = n + 1
                vals(n) = Abs(CDbl(c.Value))
            End If
        End If
    Next c
    If n = 0 Then Exit Sub

    For i = 1 To n - 1
        For k = i + 1 To n
            If vals(k) < vals(i) Then
                t = vals(i): vals(i) = vals(k): vals(k) = t
            End If
        Next k
    Next i
    i = CLng(Int(0.9 * n))
    If i < 1 Then i = 1
    If i > n Then i = n
    M = vals(i)
    If M <= 0 Then M = vals(n)
    If M <= 0 Then Exit Sub

    Set fc = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
    With fc.ColorScaleCriteria(1)
        .Type = xlConditionValueNumber
        .Value = -M
        .FormatColor.Color = RGB(230, 124, 115)
    End With
    With fc.ColorScaleCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .FormatColor.Color = RGB(255, 255, 255)
    End With
    With fc.ColorScaleCriteria(3)
        .Type = xlConditionValueNumber
        .Value = M
        .FormatColor.Color = RGB(125, 191, 138)
    End With
End Sub
