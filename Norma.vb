Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Windows

Public Module Norma

    <CommandMethod("SugereQuantidadePontos")>
    Public Sub SugereQuantidadePontos()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            Dim pl As Polyline = CType(trans.GetObject(RetornaPerimetro(ed).ObjectId, OpenMode.ForRead), Polyline)

            Dim area As Double = (pl.Area) / 10000 'Área englobada pela polyline em metros²
            Dim perimetro As Double = (pl.Length) / 100 'Perímetro da polyline em metros

            MsgBox(area.ToString + vbNewLine + perimetro.ToString)

            Dim sugestaoTexto As String

            sugestaoTexto = SugereTomadas(perimetro)
            sugestaoTexto += vbNewLine + SugereLuminarias(area)

            DesenhaMText(sugestaoTexto)

            trans.Commit()

        End Using

    End Sub

    Private Function SugereLuminarias(area As Double)

        Dim potencia As Integer

        If area <= 6 Then

            potencia = 100

        Else

            potencia = 100 + Math.Floor(((area - 6) / 4)) * 60

        End If

        Return ("Iluminação: " + potencia.ToString + " VA")

    End Function

    Private Function SugereTomadas(perimetro As Double)

        Dim qtdTomadas

        If PerguntaSeAmbienteSeco = True Then

            qtdTomadas = Math.Ceiling(perimetro / 5)
            Return ("Tomadas: " + qtdTomadas.ToString + "x100 VA")

        Else

            If PerguntaSeCozinhaOuAnalogo() = True Then

                Dim obs As String = vbNewLine + "OBS: Mínimo 2 tomadas sobre a pia."

                If perimetro <= 7 Then

                    qtdTomadas = 2
                    Return ("Tomadas: " + qtdTomadas.ToString + "x600 VA") + obs

                Else

                    qtdTomadas = Math.Ceiling(perimetro / 3.5)

                    If qtdTomadas <= 3 Then

                        Return ("Tomadas: " + qtdTomadas.ToString + "x600 VA") + obs

                    ElseIf qtdTomadas <= 6 Then

                        Return ("Tomadas: 3x600 VA + " + (qtdTomadas - 3).ToString + "x100 VA") + obs

                    Else

                        Return ("Tomadas: 2x600 VA + " + (qtdTomadas - 2).ToString + "x100 VA") + obs

                    End If

                End If

            Else

                qtdTomadas = 1
                Return ("Tomadas: " + qtdTomadas.ToString + "x600 VA")

            End If

        End If

    End Function

    Private Function PerguntaSeAmbienteSeco()

        Dim ambienteSeco As Boolean = True

        Dim opcoes() As String = {"Ambiente 'seco' (salas, quartos, corredores e afins)", "Ambiente 'molhado' (cozinhas, áreas de serviço, banheiros e afins."}
        Dim respostaUsuario As String = PedeInputUsuario(opcoes)
        If respostaUsuario = "1" Then
            ambienteSeco = True
        ElseIf respostaUsuario = "2" Then
            ambienteSeco = False
        Else
            MsgBox("Erro na escolha do tipo de ambiente! Ambiente será considerado 'seco'.")
        End If

        Return ambienteSeco

    End Function


    Private Function PerguntaSeCozinhaOuAnalogo()

        Dim cozinhaOuAnalogo As Boolean = True

        Dim opcoes() As String = {"Cozinha, copa, copa-cozinha, área de serviço, cozinha-área de serviço, lavanderia ou local análogo.", "Banheiro ou varanda."}
        Dim respostaUsuario As String = PedeInputUsuario(opcoes)
        If respostaUsuario = "1" Then
            cozinhaOuAnalogo = True
        ElseIf respostaUsuario = "2" Then
            cozinhaOuAnalogo = False
        Else
            MsgBox("Erro na escolha do tipo de ambiente! Ambiente será considerado cozinha, copa, copa-cozinha, área de serviço, cozinha-área de serviço, lavanderia ou local análogo.")
        End If

        Return cozinhaOuAnalogo

    End Function

End Module
