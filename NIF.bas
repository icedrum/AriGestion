Attribute VB_Name = "NIF"
 

Public Function Comprobar_NIF(NIF As String) As Boolean

    '-- Comprobación general de NIF

    If Len(NIF) <> 9 Then

        Comprobar_NIF = False

        Exit Function

    Else

        If IsNumeric(Mid(NIF, 1, 1)) Then

            '-- Comienza por número

            If IsNumeric(Mid(NIF, 9, 1)) Then

                Comprobar_NIF = False

                Exit Function

            Else

                Comprobar_NIF = Comprobar_NIF_PersonaFisica(NIF)

            End If

        Else

            '-- comienza por letra

            If IsNumeric(Mid(NIF, 9, 1)) Then

                '-- Acaba en número

                If InStr(1, "ABCDEFGHJPQSNU", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es una sociedad

                    Comprobar_NIF = Comprobar_NIF_Sociedad(NIF)

                ElseIf InStr(1, "T", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es un NIF antiguo que no lleva comprobación

                    Comprobar_NIF = True

                End If

            Else

                '-- Acaba en letra

                If InStr(1, "ABCDEFGHJPQSN", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es una sociedad

                    Comprobar_NIF = Comprobar_NIF_Sociedad(NIF)

                ElseIf InStr(1, "XY", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es un extranjero

                    Comprobar_NIF = Comprobar_NIF_PersonaExtranjera(NIF)

                ElseIf InStr(1, "KL", Mid(NIF, 1, 1)) <> 0 Then

                    '-- Es un NIF antiguo que no lleva comprobación

                    Comprobar_NIF = True

                End If

            End If

        End If

    End If

End Function

Public Function Comprobar_NIF_PersonaFisica(NIF As String) As Boolean

    Dim mCadena As String

    Dim mLetra As String

    Dim m23 As Integer

    mCadena = "TRWAGMYFPDXBNJZSQVHLCKE"

    '-- Tomamos el NIF propiamente dicho y calculamos el módulo 23

    m23 = Val(Mid(NIF, 1, 8)) Mod 23

    mLetra = Mid(mCadena, m23 + 1, 1)

    '-- Validamos que la letra es correcta

    If Mid(NIF, 9, 1) = mLetra Then

        Comprobar_NIF_PersonaFisica = True

    Else

        Comprobar_NIF_PersonaFisica = False

    End If

End Function

 

Public Function Comprobar_NIF_PersonaExtranjera(NIF As String) As Boolean

    Dim mCadena As String

    Dim mLetra As String

    Dim m23 As Integer

    'mCadena = "DTRWAGMYFPXBNJZSQVHLCKE"  antes enero 2012
    mCadena = "TRWAGMYFPDXBNJZSQVHLCKE"

    '-- Tomamos el NIF propiamente dicho y calculamos el módulo 23

    m23 = Val(Mid(NIF, 2, 7)) Mod 23

    mLetra = Mid(mCadena, m23 + 1, 1)

    '-- Validamos que la letra es correcta

    If Mid(NIF, 9, 1) = mLetra Then

        Comprobar_NIF_PersonaExtranjera = True

    Else

        Comprobar_NIF_PersonaExtranjera = False

    End If

End Function

 

Public Function Comprobar_NIF_Sociedad(NIF As String) As Boolean

    Dim mCadena As String

    Dim mLetra As String

    Dim vNif As String

    Dim mN2 As String

    Dim I, I2 As Integer

    Dim Suma, Control As Long

    mCadena = "ABCDEFGHIJ"
   
    vNif = Mid(NIF, 2, 7)

    '-- Sumamos las cifras pares

    For I = 2 To Len(vNif) Step 2

        Suma = Suma + Val(Mid(vNif, I, 1))

    Next I

    '-- Ahora las impares * 2, y sumando las cifras del resultado

    For I = 1 To Len(vNif) Step 2

        mN2 = CStr(Val(Mid(vNif, I, 1)) * 2)

        For I2 = 1 To Len(mN2)

            Suma = Suma + Val(Mid(mN2, I2, 1))

        Next I2

    Next I

    '-- Ya tenemos la suma y calculamos el control

    Control = 10 - (Suma Mod 10)

    If Control = 10 Then Control = 0

    mLetra = Mid(NIF, 9, 1)

    If IsNumeric(mLetra) Then

        If Val(mLetra) = Control Then

            Comprobar_NIF_Sociedad = True

        Else

            Comprobar_NIF_Sociedad = False

        End If

    Else

        If Control = 0 Then Control = 10

        If mLetra = Chr(64 + Control) Then

            Comprobar_NIF_Sociedad = True

        Else

            Comprobar_NIF_Sociedad = False

        End If

    End If

End Function

