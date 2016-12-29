Attribute VB_Name = "modRecaudacionEjecutiva"
Option Explicit


'******************************************************************************
'******************************************************************************
'******************************************************************************
'
'  Todo esta basado en un Fichero que nos pasaron (ver JIRA)
'
'******************************************************************************
'******************************************************************************
'******************************************************************************

Private NomFichHacienda As String

Dim SQL As String
Dim NF As Integer
Dim RS As ADODB.Recordset

Public Function GeneraFicheroRecaudacionEjecutiva(SQLVtos As String) As Boolean
Dim Borrar As Boolean
        
    On Error GoTo eGeneraFicheroRecaudacionEjecutiva
    
    GeneraFicheroRecaudacionEjecutiva = False
    FijarNombreFicheroTributacion
    
    'Lo generara en local para luego copiarlo donde digan
    If Dir(App.Path & "\" & NomFichHacienda, vbArchive) <> "" Then
        SQL = "Ya existe el fichero. �Reemplazar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    'Generarremos el fichero con los datos incluids
    If GeneraFichero(SQLVtos) Then
        'Marcamos con fecha de recaudacion ejecutiva
        Do
            SQL = GetFolder("Recaudacion ejecutiva")
            If SQL = "" Then
                If MsgBox("Desea cancelar el proceso", vbQuestion + vbYesNo) = vbYes Then
                    SQL = "!" 'para cancelar
                End If
            End If
        Loop Until SQL <> ""
        
        If SQL = "!" Then
            Borrar = True
        Else
            If CopiarArchivo Then
                'ACtualixar
                SQL = "UPDATE scobro SET fecejecutiva='" & Format(Now, FormatoFecha) & "'"
                SQL = SQL & " WHERE (numserie,codfaccl,fecfaccl,numorden) IN (" & SQLVtos & ")"
                EjecutarSQL SQL
                
                GeneraFicheroRecaudacionEjecutiva = True
                
            Else
                Borrar = True
            End If
        
        End If
        If Borrar Then Kill App.Path & "\" & NomFichHacienda
    End If
    
    Exit Function
eGeneraFicheroRecaudacionEjecutiva:
    MuestraError Err.Number, Err.Description
End Function



Private Function GeneraFichero(SQLVtos) As Boolean
Dim Contador As Integer
Dim Impor As Currency


    On Error GoTo eGeneraFichero

    GeneraFichero = False
    
    NF = FreeFile
    SQL = ""
    
    Open App.Path & "\" & NomFichHacienda For Output As #NF
    
    'Si ha llegado aqui el fichero esta abierto
    Contador = 0
    SQL = "Select numserie,codfaccl,fecfaccl,numorden,fecvenci,impvenci,gastos,impcobro,scobro.codmacta,nommacta,nifdatos"
    SQL = SQL & ",dirdatos,codposta,despobla,desprovi,codbanco ,codsucur,digcontr,scobro.cuentaba"
    SQL = SQL & ",text33csb,text41csb,text42csb,text43csb,text51csb,text52csb,text53csb,text61csb,text62csb,text63csb,text71csb,text72csb,text73csb,text81csb,text82csb,text83csb"
    SQL = SQL & " FROM scobro,cuentas  WHERE scobro.codmacta=cuentas.codmacta "
    SQL = SQL & " AND (numserie,codfaccl,fecfaccl,numorden) IN (" & SQLVtos & ")"
    SQL = SQL & " ORDER BY numserie,codfaccl,numorden"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText  'keyset para hacer forward
    Impor = 0
    While Not RS.EOF
        'Voy guardandome el valor para el final
        Contador = Contador + 1
        If Not IsNull(RS!Gastos) Then Impor = Impor + RS!Gastos
        If Not IsNull(RS!impcobro) Then Impor = Impor - RS!impcobro
        Impor = Impor + RS!impvenci
       
        'Imprime registros 51,52,53,55
        L51
        L52
        L53
        L55
        'Registros de texto
        Textos

        RS.MoveNext
    Wend
    RS.Close
    Totales Contador, Impor
    GeneraFichero = True
    
eGeneraFichero:
    If Err.Number <> 0 Then MuestraError Err.Number
    CerrarFichero
    Set RS = Nothing
        
            
End Function


Private Sub CerrarFichero()
    On Error Resume Next
    Close #NF
    Err.Clear
End Sub


'RS es variable globarl del  modulo
Private Sub L51()
Dim Impor As Currency
'    1- 2   2 N Tipo Registro: Contenido 51 .
'    3- 4   2 N Anualidad.
'    5- 7   3 N C�digo Ayuntamiento (u Organismo), seg�n Anexo II.
'    8-10   3 N C�digo Tributo, seg�n Anexo III.
'    11-11  1 N Tipo Valor. Valores posibles: 1, 2, 3, o 4.
'    12-18  7 X Referencia Emisor.
'    19-38  20 X N�mero Fijo.
'    39-39  1 X Personalidad: F-�sica, J-ur�dica, E-xtranjero, o I-ncorrecto.
'    40-48  9 X N.I.F. / C.I.F. (Seg�n Normas Codificaci�n Anexo VI).
'    49-50  2 N N�. de Subconceptos que posee el Valor.
'    51-59  9 N Importe Total Euros (7 enteros y 2 decimales sin coma)
'    60-68  9 N Importe Principal Euros (7 enteros y 2 decimales sin coma).
'    69-76  8 N Recargo Aplicado Euros (6 enteros y 2 decimales sin coma)
'    77-82  6 X Fecha Notificaci�n (ddmmaa).
    'EJEMPLo
    '51083427783  00081000081              F20715635X0100000286800000286800000000000000
    '51083427783  02424002424              F20776902J0100000573600000573600000000000000

    SQL = "51"
    SQL = SQL & Right(Year(Now), 2) 'anualidad
    SQL = SQL & "342" 'cod eayto u organismo
    SQL = SQL & "778" 'cod del tributo
    SQL = SQL & "3" ' Tributos de car�cter peri�dico en ejecutiva 3
    
    SQL = SQL & Right(Format(RS!codfaccl, "0000000"), 7) 'referencia emisor
    SQL = SQL & Left(RS!NUmSerie & "   ", 3) & Format(RS!codfaccl, "0000000") & " " & Format(RS!numorden, "000") & Format(RS!fecfaccl, "ddmmyy")  'numero fijo UNICO para el vtot
    
    'Si la primera es un numero es persona fisica, si no juridica
    If IsNumeric(Mid(RS!nifdatos, 1, 1)) Then
        SQL = SQL & "F" 'persona fisica
    Else
        SQL = SQL & "J" 'persona fisica
    End If
    SQL = SQL & Mid(RS!nifdatos & "   ", 1, 9)
    SQL = SQL & "00" 'numero subconceptos =1
    Impor = RS!impvenci + DBLet(RS!Gastos, "N") - DBLet(RS!impcobro, "N")
   
    SQL = SQL & Right(String(9, "0") & CStr(Val(Impor * 100)), 9)
    SQL = SQL & Right(String(9, "0") & CStr(Val(Impor * 100)), 9)
    SQL = SQL & String(8, "0") 'recargo
    SQL = SQL & String(6, "0") 'fecha noticifacion. En el ejmplo estaba a cero
    Print #NF, SQL
    
End Sub


Private Sub L52()
'    1- 2 2 N Tipo Registro: Contenido 52.
'    3-42 40 X Apellido_1� Apellido_2�, Nombre.
'    43-44 2 X Siglas V�a, seg�n tabla del ANEXO IV.
'    45-82 38 X Domicilio Fiscal.
    '52ALVENTOSA CASANOVA MIQUEL                 AVDA LUIS SUNER 19,6-16a
    
  '  dirdatos ,
    SQL = "52"
    SQL = SQL & Mid(RS!Nommacta & Space(40), 1, 40)
    SQL = SQL & Right(Space(40) & RS!dirdatos, 38)
    Print #NF, SQL
End Sub

Private Sub L53()
'    1- 2 2 N Tipo Registro: Contenido 53.
'    3- 4 2 N C�digo del I.N.E. de la Provincia del Domicilio Fiscal.
'    5- 7 3 N C�digo del I.N.E. del Municipio del Domicilio Fiscal.
'    8-12 5 X C�digo Postal.
'    13-37 25 X Municipio (Literal).
'    38-82 45 X Objeto Tributario (Domicilio Tributario, Parcela, Matr�cula).
    '530000046600ALZIRA                              0000000000000000000000000000000000
    SQL = "53"
    SQL = SQL & "00000"
    SQL = SQL & Right(Space(5) & RS!codposta, 5)
    SQL = SQL & Mid(RS!despobla & Space(70), 1, 70)
    Print #NF, SQL
End Sub


Private Sub L55()
'    1- 2 2 N Tipo Registro: Contenido 55.
'    3- 6 4 N C�digo de la Entidad Bancaria de Cobro, seg�n C.S.I.
'    7-10 4 N C�digo de la Surcursal de la Entidad Bancaria, seg�n C.S.I.
'    11-12 2 X D�gitos de Control de la Cuenta.
'    13-22 10 N N�mero de Cuenta.
'    23-62 40 X Titular de la Cuenta (Apellido_1� Apellido_2�, Nombre).
'    63-63 1 X Personalidad: F-�sica, J-ur�dica, E-xtranjero, o I-ncorrecto.
'    64-72 9 X N.I.F. / C.I.F. (Seg�n Normas Codificaci�n Anexo VI).
'    73-82 10 X NO_DEFINIDO (relleno a espacios en blanco).
    '5530821308888888888888ALVENTOSA CASANOVA MIQUEL               F20715635X0000000000
    ',codbanco ,codsucur,digcontr,scobro.cuentab
    SQL = "55"
    SQL = SQL & Format(RS!codbanco, "0000") & Format(RS!codsucur, "0000")
    SQL = SQL & RS!digcontr & Right(String(10, "0") & RS.Fields("cuentaba"), 10)
    SQL = SQL & Mid(RS!Nommacta & Space(40), 1, 40)
    
    'Si la primera es un numero es persona fisica, si no juridica
    If IsNumeric(Mid(RS!nifdatos, 1, 1)) Then
        SQL = SQL & "F" 'persona fisica
    Else
        SQL = SQL & "J" 'persona fisica
    End If
    
    SQL = SQL & Mid(RS!nifdatos & "   ", 1, 9)
    
    SQL = SQL & String(10, " ")
    Print #NF, SQL
End Sub



Private Sub Textos()
Dim I As Byte
Dim Lin As Byte
Dim J As Byte
Dim Aux As String
    '1- 2 2 N Tipo de Registro: Contenido 61..66.
    '3-82 80 X L�nea de Texto.
'    61talla08-no reg
'    62Ordinaria 9,07 eur/fan
'    63PLA            8  97      2  66   ,0286
'    64PLA            8 137 1    1       ,0286
'    65CASTELLET     11  79      3       ,0286
'    66    TOTAL . . . . . .     6  66
'


    'text33csb,text41csb,text42csb,text43csb,text51csb,text52csb,text53csb,text61csb,
    'text62csb,text63csb,text71csb,text72csb,text73csb,text81csb,text82csb,text83csb
    
    Lin = 61
    For I = 0 To 5          'son, como mucho, 6 lineas
        'Empieza = 19 'Empieza en el txt33 es el campo(23)
        
        'De dos en dos
        J = I * 2
        J = 19 + J
        'Debug.Print RS.Fields(J + 0).Name
        
        Aux = Mid(DBLet(RS.Fields(J + 0), "T") & Space(40), 1, 40)
        Aux = Aux & Mid(DBLet(RS.Fields(J + 1), "T") & Space(40), 1, 40)
        If Trim(Aux) <> "" Then
            SQL = CStr(Lin) & Aux
            Print #NF, SQL
            Lin = Lin + 1
        End If
    Next
    
End Sub



Private Sub Totales(Registros As Integer, Importe As Currency)
'        1- 2 2 N Tipo de Registro: Contenido 98.
'        3-10 8 N N�mero de Valores.
'        11-22 12 N Importe Total de los Valores (10 enteros y 2 decimales sin coma)
'        23-25 3 N N�mero Total de Subconceptos Distintos.
'        26-31 6 X Fecha Generaci�n (ddmmaa).
'        32-42 11 X Nombre del Fichero (Tejaaavpttt - sin el �.�-), ver apartado 4.2.
'        43-82 40 X NO_DEFINIDO (relleno a espacios en blanco).

'   98000000080000079401   120701T0134230           0000000000000000000000000000000000
    SQL = "98"
    SQL = SQL & Format(Registros, "00000000")
    
    SQL = SQL & Right(String(12, "0") & Val(Importe * 100), 12)
    SQL = SQL & "000"
    SQL = SQL & Format(Now, "ddmmyy")
    SQL = SQL & Mid(NomFichHacienda, 1, 8) & "   "
    SQL = SQL & Space(40)
    SQL = Mid(SQL, 1, 82)
    Print #NF, SQL
End Sub



Private Sub FijarNombreFicheroTributacion()

' : Tejaaavp , Donde:
'   ej      son los dos �ltimos d�gitos del a�o en que se quieren poner al cobro los
'           alores (admiti�ndose �nicamente el a�o corriente, y el a�o siguiente durante
'           el �ltimo trimestre del corriente).

'   aaa     es el c�digo num�rico (de 3 d�gitos) correspondiente al Ayuntamiento (u
'           Organismo), seg�n la codificaci�n de uso habitual por la Diputaci�n (Anexo
'           II).
'   v       es el c�digo del tipo de valor de los valores que se env�an, normalmente ser�
'           1 (Recibos en Voluntaria), pero, tambi�n pueden ser 2, 3 o 4.
'   p es un d�gito del 0 al 9. El d�gito 0 se utilizar� con aquellos tributos que
'           tengan car�cter anual, es decir, del mismo tributo s�lo se produce una
'           emisi�n durante el ejercicio. Los d�gitos del 1 al 9 ser�n utilizados para los
'           distintos ficheros que se generen en cada una de las emisiones peri�dicas del
'           mismo tributo en un mismo ejercicio, o, dentro de un mismo tributo, para
'           diferenciar entre distintos paquetes de valores (pero que no se correspondan
'           con emisiones peri�dicas, o no se puedan distinguir mediantes subconceptos
'           tributarios). La numeraci�n ser� correlativa dentro del mismo a�o de cargo,
'           empezando por 1. Por ejemplo, considerar un municipio con dos padrones de
'           Basura, uno el del n�cleo y otro el del diseminados, el primero tendr�a como
'           valor p un 1, y el segundo un 2.
' 3�. La segunda parte del nombre tendr� la estructura: ttt.
'   ttt     ser� el c�digo num�rico (de 3 d�gitos) correspondiente al concepto del
'           tributo, seg�n la tabla del Anexo III.
    'Linea 51
    '    SQL = SQL & "342" 'cod eayto u organismo
    '    SQL = SQL & "778" 'cod del tributo
    '    SQL = SQL & "3" ' Tributos de car�cter peri�dico en ejecutiva 3
    If Month(Now) > 8 Then
        NomFichHacienda = Year(Now) + 1
    Else
        NomFichHacienda = Year(Now)
    End If
    NomFichHacienda = Right(NomFichHacienda, 2) 'dos ult del a�o
    NomFichHacienda = "T" & NomFichHacienda
    NomFichHacienda = NomFichHacienda & "342"
    NomFichHacienda = NomFichHacienda & "3"
    NomFichHacienda = NomFichHacienda & "0"
    NomFichHacienda = NomFichHacienda & "." & "778"
End Sub



Private Function CopiarArchivo() As Boolean

    FileCopy App.Path & "\" & NomFichHacienda, SQL & "\" & NomFichHacienda
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        CopiarArchivo = False
    Else
        CopiarArchivo = True
    End If
        
End Function
