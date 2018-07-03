Option Explicit On
Option Strict On


Public Class Form1
    Private TextoIncorrecto As Boolean


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim NumeroResultado As Long, CadenaDeTextoResultado As String
        TextoIncorrecto = False
        TextBox1.Text = TextBox1.Text.ToLower
        If TextBox1.TextLength = 0 Then
            AdvertimosAlUsuario("Tienes que introducir algo en el cuadro de texto", "Cuadro de texto vacío")
        ElseIf EsUnNumeroEntero(TextBox1.Text) Then
            If TextBox1.TextLength > 12 Then
                AdvertimosAlUsuario("No olvides que, si quieres introducir un número, tienes que ser un número NO negativo menor que 1 BILLÓN", "Número inválido")
            Else
                lblNumerico.Text = FormatNumber(TextBox1.Text, 0)
                CadenaDeTextoResultado = DeNumeroACastellano(CType(TextBox1.Text, Long))
                lblTexto.Text = CadenaDeTextoResultado
            End If
        ElseIf EsUnaCadenaDeTexto(TextBox1.Text) Then
            NumeroResultado = DeCastellanoANumero(TextBox1.Text)
            If TextoIncorrecto Then
                AdvertimosAlUsuario("La cadena de texto que has introducido no es válida", "Cadena de texto inválida")
            Else
                lblNumerico.Text = FormatNumber(NumeroResultado, 0)
                lblTexto.Text = TextBox1.Text
            End If
        Else
            AdvertimosAlUsuario("Mira a ver qué has hecho, porque lo que has introducido en el cuadro de texto no es un número entero ni tampoco una cadena de texto válida")
        End If




    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = vbCr Then Button1.PerformClick()
    End Sub



    Function DeNumeroACastellano(ByRef Numero As Long) As String
        Dim Millardos, Millones, Miles, Unidades As Long
        ' Se entiende que las unidades son por debajo de 1.000, los miles son por debajo del millón, y los millones son por debajo del millardo
        Dim IzquierdaAuxiliar, DerechaAuxiliar As String

        Unidades = Numero Mod 1000
        Miles = (Numero \ 1000) Mod 1000
        Millones = (Numero \ 1000000) Mod 1000
        Millardos = Numero \ 1000000000

        If Numero < 0 Then
            CortamosConError("Recuerda que hay que meter un número NO NEGATIVO menor que un millón") : Stop : End
        ElseIf Numero < 1000 Then
            Select Case Numero
                Case 0 : Return "cero"
                Case 1 : Return "uno"
                Case 2 : Return "dos"
                Case 3 : Return "tres"
                Case 4 : Return "cuatro"
                Case 5 : Return "cinco"
                Case 6 : Return "seis"
                Case 7 : Return "siete"
                Case 8 : Return "ocho"
                Case 9 : Return "nueve"
                Case 10 : Return "diez"
                Case 11 : Return "once"
                Case 12 : Return "doce"
                Case 13 : Return "trece"
                Case 14 : Return "catorce"
                Case 15 : Return "quince"

                Case 16 : Return "dieciséis"
                Case 17 To 19 : Return "dieci" & DeNumeroACastellano(Numero Mod 10)

                Case 20 : Return "veinte"
                Case 22 : Return "veintidós"
                Case 23 : Return "veintitrés"
                Case 26 : Return "veintiséis"
                Case 21 To 29 : Return "veinti" & DeNumeroACastellano(Numero Mod 10)

                Case 30 : Return "treinta"
                Case 40 : Return "cuarenta"
                Case 50 : Return "cincuenta"
                Case 60 : Return "sesenta"
                Case 70 : Return "setenta"
                Case 80 : Return "ochenta"
                Case 90 : Return "noventa"

                Case 31 To 99 : Return DeNumeroACastellano(10 * (Numero \ 10)) & " y " & DeNumeroACastellano(Numero Mod 10)

                Case 100 : Return "cien"
                Case 101 To 199 : Return "ciento " & DeNumeroACastellano(Numero Mod 100)

                Case 500 : Return "quinientos"
                Case 700 : Return "setecientos"
                Case 900 : Return "novecientos"
                Case 200, 300, 400, 600, 800 : Return DeNumeroACastellano(Numero \ 100) & "cientos"
                Case 201 To 999 : Return DeNumeroACastellano(100 * (Numero \ 100)) & " " & DeNumeroACastellano(Numero Mod 100)
            End Select

        ElseIf Numero < 1000000 Then

            If Unidades = 0 Then DerechaAuxiliar = vbNullString Else DerechaAuxiliar = " " & DeNumeroACastellano(Unidades)
            IzquierdaAuxiliar = DeNumeroACastellano(Miles)

            If Miles = 1 Then
                Return "mil" & DerechaAuxiliar
            ElseIf Miles Mod 10 = 1 And Miles Mod 100 <> 11 Then
                If Not IzquierdaAuxiliar.EndsWith("uno") Then CortamosConError("No deberíamos estar aquí") : Stop : End
                IzquierdaAuxiliar = IzquierdaAuxiliar.Substring(0, IzquierdaAuxiliar.Length - 3)
                If Miles Mod 100 = 21 Then
                    IzquierdaAuxiliar = IzquierdaAuxiliar & "ún"
                Else
                    IzquierdaAuxiliar = IzquierdaAuxiliar & "un"
                End If
            Else
                If IzquierdaAuxiliar.EndsWith("uno") Then CortamosConError("No deberíamos estar aquí") : Stop : End
            End If
            Return IzquierdaAuxiliar & " mil" & DerechaAuxiliar

        ElseIf Numero < 1000000000 Then
            If Unidades = 0 And Miles = 0 Then DerechaAuxiliar = vbNullString Else DerechaAuxiliar = " " & DeNumeroACastellano(1000 * Miles + Unidades)
            IzquierdaAuxiliar = DeNumeroACastellano(Millones)

            If Millones = 1 Then
                Return "un millón" & DerechaAuxiliar
            ElseIf Millones Mod 10 = 1 And Millones Mod 100 <> 11 Then
                If Not IzquierdaAuxiliar.EndsWith("uno") Then CortamosConError("No deberíamos estar aquí") : Stop : End
                IzquierdaAuxiliar = IzquierdaAuxiliar.Substring(0, IzquierdaAuxiliar.Length - 3)
                If Millones Mod 100 = 21 Then
                    IzquierdaAuxiliar = IzquierdaAuxiliar & "ún"
                Else
                    IzquierdaAuxiliar = IzquierdaAuxiliar & "un"
                End If
            Else
                If IzquierdaAuxiliar.EndsWith("uno") Then CortamosConError("No deberíamos estar aquí") : Stop : End
            End If
            Return IzquierdaAuxiliar & " millones" & DerechaAuxiliar

        ElseIf Numero < 1000000000000 Then
            If Millones = 0 Then
                If Miles = 0 And Unidades = 0 Then DerechaAuxiliar = " millones" Else DerechaAuxiliar = " millones " & DeNumeroACastellano(1000 * Miles + Unidades)
            ElseIf Millones = 1 Then
                If Miles = 0 And Unidades = 0 Then DerechaAuxiliar = " un millones" Else DerechaAuxiliar = " un millones " & DeNumeroACastellano(1000 * Miles + Unidades)

            Else
                If Miles = 0 And Unidades = 0 Then
                    DerechaAuxiliar = " " & DeNumeroACastellano(1000000 * Millones)
                Else
                    DerechaAuxiliar = " " & DeNumeroACastellano(1000000 * Millones) & " " & DeNumeroACastellano(1000 * Miles + Unidades)
                End If
            End If

            IzquierdaAuxiliar = DeNumeroACastellano(Millardos)

            If Millardos = 1 Then
                Return "mil" & DerechaAuxiliar
            ElseIf Millardos Mod 10 = 1 And Millardos Mod 100 <> 11 Then
                If Not IzquierdaAuxiliar.EndsWith("uno") Then CortamosConError("No deberíamos estar aquí") : Stop : End
                IzquierdaAuxiliar = IzquierdaAuxiliar.Substring(0, IzquierdaAuxiliar.Length - 3)
                If Millardos Mod 100 = 21 Then
                    IzquierdaAuxiliar = IzquierdaAuxiliar & "ún"
                Else
                    IzquierdaAuxiliar = IzquierdaAuxiliar & "un"
                End If
            Else
                If IzquierdaAuxiliar.EndsWith("uno") Then CortamosConError("No deberíamos estar aquí") : Stop : End
            End If
            Return IzquierdaAuxiliar & " mil" & DerechaAuxiliar
        Else
            CortamosConError("Recuerda que hay que meter un número no negativo menor que UN BILLÓN") : Stop : End
        End If
    End Function

    Function DeCastellanoANumero(NumeroEnLetras As String) As Long
        Dim TextoIzquierdaAuxiliar, TextoDerechaAuxiliar As String
        Dim NumeroIzquierdaAuxiliar, NumeroDerechaAuxiliar As Long

        If NumeroEnLetras <> "cero" And NumeroEnLetras.IndexOf("cero") <> -1 Then TextoIncorrecto = True : Return -1
        If NumeroEnLetras = vbNullString Then TextoIncorrecto = True : Return -1

        Select Case CuantasApariciones("mill", NumeroEnLetras)
            Case 1
                If NumeroEnLetras.IndexOf("un millón") <> -1 Then
                    TextoDerechaAuxiliar = NumeroEnLetras.Substring(NumeroEnLetras.IndexOf("un millón") + 9)
                    If TextoDerechaAuxiliar = vbNullString Then
                        NumeroDerechaAuxiliar = 0
                    ElseIf TextoDerechaAuxiliar.StartsWith(" ") Then
                        TextoDerechaAuxiliar = TextoDerechaAuxiliar.Substring(1)
                        If TextoDerechaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
                        NumeroDerechaAuxiliar = DeCastellanoANumero(TextoDerechaAuxiliar)
                    Else
                        TextoIncorrecto = True : Return -1
                    End If
                    Return 1000000 + NumeroDerechaAuxiliar
                ElseIf NumeroEnLetras.IndexOf(" millones") <> -1 Then
                    TextoDerechaAuxiliar = NumeroEnLetras.Substring(NumeroEnLetras.IndexOf(" millones") + 9)
                    If TextoDerechaAuxiliar = vbNullString Then
                        NumeroDerechaAuxiliar = 0
                    ElseIf TextoDerechaAuxiliar.StartsWith(" ") Then
                        TextoDerechaAuxiliar = TextoDerechaAuxiliar.Substring(1)
                        If TextoDerechaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
                        NumeroDerechaAuxiliar = DeCastellanoANumero(TextoDerechaAuxiliar)
                    Else
                        TextoIncorrecto = True : Return -1
                    End If

                    TextoIzquierdaAuxiliar = NumeroEnLetras.Substring(0, NumeroEnLetras.IndexOf(" millones"))
                    If TextoIzquierdaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
                    If TextoIzquierdaAuxiliar.EndsWith("uno") Then
                        TextoIncorrecto = True : Return -1
                    ElseIf TextoIzquierdaAuxiliar.EndsWith("veintiún") Then
                        TextoIzquierdaAuxiliar = TextoIzquierdaAuxiliar.Substring(0, TextoIzquierdaAuxiliar.Length - 8) & "veintiuno"
                    ElseIf TextoIzquierdaAuxiliar.EndsWith(" un") Then
                        TextoIzquierdaAuxiliar = TextoIzquierdaAuxiliar.Substring(0, TextoIzquierdaAuxiliar.Length - 3) & " uno"
                    End If
                    NumeroIzquierdaAuxiliar = DeCastellanoANumero(TextoIzquierdaAuxiliar)
                    Return 1000000 * NumeroIzquierdaAuxiliar + NumeroDerechaAuxiliar

                Else
                    TextoIncorrecto = True : Return -1
                End If
            Case 0
            Case Else : TextoIncorrecto = True : Return -1
        End Select

        Select Case CuantasApariciones("mil", NumeroEnLetras)
            Case 1
                If NumeroEnLetras.StartsWith("mil") Then
                    If NumeroEnLetras = "mil" Then
                        Return 1000
                    Else
                        If Not NumeroEnLetras.StartsWith("mil ") Then TextoIncorrecto = True : Return -1
                        TextoDerechaAuxiliar = NumeroEnLetras.Substring(4)
                        If TextoDerechaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
                        Return 1000 + DeCastellanoANumero(TextoDerechaAuxiliar)
                    End If
                ElseIf NumeroEnLetras.IndexOf(" mil") <> -1 Then
                    TextoDerechaAuxiliar = NumeroEnLetras.Substring(NumeroEnLetras.IndexOf(" mil") + 4)
                    If TextoDerechaAuxiliar.StartsWith(" ") Then
                        TextoDerechaAuxiliar = TextoDerechaAuxiliar.Substring(1)
                        If TextoDerechaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
                    End If

                    TextoIzquierdaAuxiliar = NumeroEnLetras.Substring(0, NumeroEnLetras.IndexOf(" mil"))
                    If TextoIzquierdaAuxiliar.EndsWith("uno") Then
                        TextoIncorrecto = True : Return -1
                    ElseIf TextoIzquierdaAuxiliar.EndsWith("veintiún") Then
                        TextoIzquierdaAuxiliar = TextoIzquierdaAuxiliar.Substring(0, TextoIzquierdaAuxiliar.Length - 8) & "veintiuno"
                    ElseIf TextoIzquierdaAuxiliar.EndsWith(" un") Then
                        TextoIzquierdaAuxiliar = TextoIzquierdaAuxiliar.Substring(0, TextoIzquierdaAuxiliar.Length - 3) & " uno"
                    End If

                    NumeroIzquierdaAuxiliar = DeCastellanoANumero(TextoIzquierdaAuxiliar)
                    If TextoDerechaAuxiliar = vbNullString Then NumeroDerechaAuxiliar = 0 Else NumeroDerechaAuxiliar = DeCastellanoANumero(TextoDerechaAuxiliar)
                    Return 1000 * NumeroIzquierdaAuxiliar + NumeroDerechaAuxiliar
                Else
                    TextoIncorrecto = True : Return -1
                End If
            Case 0
            Case Else : TextoIncorrecto = True : Return -1
        End Select


        Select Case CuantasApariciones("ien", NumeroEnLetras)
            Case 1
                If NumeroEnLetras = "cien" Then
                    Return 100
                ElseIf NumeroEnLetras.StartsWith("ciento ") Then
                    TextoDerechaAuxiliar = NumeroEnLetras.Substring(7)
                    If TextoDerechaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
                    Return 100 + DeCastellanoANumero(TextoDerechaAuxiliar)
                End If

                If NumeroEnLetras.IndexOf("ientos") <> -1 Then
                    TextoDerechaAuxiliar = NumeroEnLetras.Substring(NumeroEnLetras.IndexOf("ientos") + 6)
                    If TextoDerechaAuxiliar = vbNullString Then
                        NumeroDerechaAuxiliar = 0
                    ElseIf TextoDerechaAuxiliar.StartsWith(" ") Then
                        TextoDerechaAuxiliar = TextoDerechaAuxiliar.Substring(1)
                        If TextoDerechaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
                        NumeroDerechaAuxiliar = DeCastellanoANumero(TextoDerechaAuxiliar)
                    Else
                        TextoIncorrecto = True : Return -1
                    End If
                End If

                If NumeroEnLetras.StartsWith("quinientos") Then
                    Return 500 + NumeroDerechaAuxiliar
                ElseIf NumeroEnLetras.StartsWith("setecientos") Then
                    Return 700 + NumeroDerechaAuxiliar
                ElseIf NumeroEnLetras.StartsWith("novecientos") Then
                    Return 900 + NumeroDerechaAuxiliar
                ElseIf NumeroEnLetras.IndexOf("cientos") <> -1 Then
                    TextoIzquierdaAuxiliar = NumeroEnLetras.Substring(0, NumeroEnLetras.IndexOf("cientos"))
                    Select Case TextoIzquierdaAuxiliar
                        Case "dos", "tres", "cuatro", "seis", "ocho"
                            NumeroIzquierdaAuxiliar = DeCastellanoANumero(TextoIzquierdaAuxiliar)
                            Return 100 * NumeroIzquierdaAuxiliar + NumeroDerechaAuxiliar
                        Case Else
                            TextoIncorrecto = True : Return -1
                    End Select
                Else
                    TextoIncorrecto = True : Return -1
                End If
            Case 0
            Case Else : TextoIncorrecto = True : Return -1
        End Select

        Select Case CuantasApariciones("nta", NumeroEnLetras)
            Case 1
                TextoDerechaAuxiliar = NumeroEnLetras.Substring(NumeroEnLetras.IndexOf("nta") + 3)
                If TextoDerechaAuxiliar = vbNullString Then
                    NumeroDerechaAuxiliar = 0
                ElseIf TextoDerechaAuxiliar.StartsWith(" y ") Then
                    TextoDerechaAuxiliar = TextoDerechaAuxiliar.Substring(3)
                    If TextoDerechaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
                    NumeroDerechaAuxiliar = DeCastellanoANumero(TextoDerechaAuxiliar)
                Else
                    TextoIncorrecto = True : Return -1
                End If

                Select Case NumeroEnLetras.Substring(0, NumeroEnLetras.IndexOf("nta"))
                    Case "trei" : Return 30 + NumeroDerechaAuxiliar
                    Case "cuare" : Return 40 + NumeroDerechaAuxiliar
                    Case "cincue" : Return 50 + NumeroDerechaAuxiliar
                    Case "sese" : Return 60 + NumeroDerechaAuxiliar
                    Case "sete" : Return 70 + NumeroDerechaAuxiliar
                    Case "oche" : Return 80 + NumeroDerechaAuxiliar
                    Case "nove" : Return 90 + NumeroDerechaAuxiliar
                    Case Else : TextoIncorrecto = True : Return -1
                End Select

            Case 0
            Case Else : TextoIncorrecto = True : Return -1
        End Select

        Select Case NumeroEnLetras
            Case "cero" : Return 0
            Case "uno" : Return 1
            Case "dos" : Return 2
            Case "tres" : Return 3
            Case "cuatro" : Return 4
            Case "cinco" : Return 5
            Case "seis" : Return 6
            Case "siete" : Return 7
            Case "ocho" : Return 8
            Case "nueve" : Return 9
            Case "diez" : Return 10
            Case "once" : Return 11
            Case "doce" : Return 12
            Case "trece" : Return 13
            Case "catorce" : Return 14
            Case "quince" : Return 15
            Case "dieciséis" : Return 16
            Case "veinte" : Return 20
            Case "veintidós" : Return 22
            Case "veintitrés" : Return 23
            Case "veintiséis" : Return 26
        End Select

        If NumeroEnLetras.StartsWith("dieci") Then
            TextoDerechaAuxiliar = NumeroEnLetras.Substring(5)
            NumeroDerechaAuxiliar = DeCastellanoANumero(TextoDerechaAuxiliar)
            Select Case NumeroDerechaAuxiliar
                Case 7, 8, 9 : Return 10 + NumeroDerechaAuxiliar
                Case Else : TextoIncorrecto = True : Return -1
            End Select
        ElseIf NumeroEnLetras.StartsWith("veinti") Then
            TextoDerechaAuxiliar = NumeroEnLetras.Substring(6)
            If TextoDerechaAuxiliar = vbNullString Then TextoIncorrecto = True : Return -1
            NumeroDerechaAuxiliar = DeCastellanoANumero(TextoDerechaAuxiliar)
            Select Case NumeroDerechaAuxiliar
                Case 1, 4, 5, 7, 8, 9 : Return 20 + NumeroDerechaAuxiliar
                Case Else : TextoIncorrecto = True : Return -1
            End Select
        Else
            TextoIncorrecto = True : Return -1
        End If

    End Function



    Private Function EsUnNumeroEntero(CadenaDeTexto As String) As Boolean
        Dim Subcadena As String
        Dim Contador As Integer
        For Contador = 0 To CadenaDeTexto.Length - 1
            Subcadena = CadenaDeTexto.Substring(Contador, 1)
            If Not Subcadena Like "#" Then Return False
        Next
        Return True
    End Function

    Private Function EsUnaCadenaDeTexto(CadenaDeTexto As String) As Boolean
        Dim Contador As Integer, Subcadena As String
        CadenaDeTexto = CadenaDeTexto.ToLower
        For Contador = 0 To CadenaDeTexto.Length - 1
            Subcadena = CadenaDeTexto.Substring(Contador, 1)
            If Not Subcadena Like "[a-z]" And Not Subcadena Like "[áéíóú]" And Not Subcadena Like " " And Not Subcadena Like "ñ" Then Return False
        Next
        Return True
    End Function

    Private Function CuantasApariciones(Subcadena As String, Cadena As String) As Integer
        Dim Pivote As Integer = 0
        Dim Acumulador As Integer = 0
        Do While Cadena.IndexOf(Subcadena, Pivote) <> -1
            Pivote = Cadena.IndexOf(Subcadena, Pivote) + 1
            Acumulador += 1
        Loop
        Return Acumulador
    End Function


    Private Sub CortamosConError(CadenaDeMensaje As String, Optional CadenaDeTitulo As String = vbNullString)
        MessageBox.Show(CadenaDeMensaje, CadenaDeTitulo, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub AdvertimosAlUsuario(CadenaDeMensaje As String, Optional CadenaDeTitulo As String = vbNullString)
        MessageBox.Show(CadenaDeMensaje, CadenaDeTitulo, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        TextBox1.SelectAll()
        TextBox1.Focus()
    End Sub



End Class
