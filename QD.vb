Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class QD

    'Private nome As String
    Private nome As String
    Private derivaDe As String
    Private esquema As String
    Private tensao As String
    Private potencia_total As Double
    Private fases As String
    Private potencia_r As Double
    Private potencia_s As Double
    Private potencia_t As Double
    Private secao As Double
    Private disjuntor As Integer
    Private barramento As Barramento
    Private circuitos As List(Of Circuito)

    Public Sub New(_nome As String, _derivaDe As String, _circuitos As List(Of Circuito))

        nome = _nome
        derivaDe = _derivaDe
        circuitos = _circuitos

        OrdenaCircuitos()
        DistribuiFases()
        SomaPotenciaTotal()
        SomaPotenciaFases()
        DefineEsquema()
        DefineFases()
        DefineTensao()
        DimensionaSecaoEDisjuntor()

        barramento = New Barramento(disjuntor)

        MsgBox(esquema.ToString + "/" + tensao.ToString + "/" + potencia_total.ToString + "/" + fases.ToString + "/" + potencia_r.ToString + "/" + potencia_s.ToString + "/" + potencia_t.ToString + "/" +
               secao.ToString + "/" + disjuntor.ToString + "/" + barramento.GetTamanho.ToString)

    End Sub

    Private Sub SomaPotenciaTotal()

        For i = 0 To (circuitos.Count - 1)
            potencia_total += circuitos(i).GetPotenciaTotal
        Next

    End Sub

    Private Sub SomaPotenciaFases()

        For i = 0 To (circuitos.Count - 1)
            potencia_r += circuitos(i).GetPotenciaR
            potencia_s += circuitos(i).GetPotenciaS
            potencia_t += circuitos(i).GetPotenciaT
        Next

    End Sub

    Private Sub DefineEsquema()

        Dim numeroDeFases = 0

        If potencia_r > 0 Then
            numeroDeFases += 1
        End If

        If potencia_s > 0 Then
            numeroDeFases += 1
        End If

        If potencia_t > 0 Then
            numeroDeFases += 1
        End If

        If numeroDeFases > 0 Then
            esquema = numeroDeFases.ToString + "F+N+T"
        Else
            MsgBox("Número de fases igual a 0 (zero)")
        End If

    End Sub

    Private Sub DefineTensao()

        If esquema = "F+N+T" Then
            tensao = "220"
        ElseIf esquema = "2F+N+T" Or esquema = "3F+N+T" Then
            tensao = "380/220"
        End If

    End Sub

    Private Sub DefineFases()

        fases = ""

        If potencia_r > 0 Then
            fases = "R"
        End If

        If potencia_s > 0 Then
            If fases = "" Then
                fases += "S"
            Else
                fases += "+S"
            End If
        End If

        If potencia_t > 0 Then
            If fases = "" Then
                fases += "T"
            Else
                fases += "+T"
            End If
        End If

    End Sub

    Private Sub DimensionaSecaoEDisjuntor()

        If potencia_total <= 0 Then

            MsgBox("Potência total <= 0")

        Else

            If esquema = "F+N+T" Then

                Select Case potencia_total
                    Case <= 0
                        Exit Select
                    Case < 6000
                        secao = 6.0
                        disjuntor = 32
                    Case < 8000
                        secao = 10.0
                        disjuntor = 40
                    Case < 11000
                        secao = 10.0
                        disjuntor = 50
                    Case < 13000
                        secao = 16.0
                        disjuntor = 63
                    Case < 15000
                        secao = 16.0
                        disjuntor = 70
                End Select


            ElseIf esquema = "2F+N+T" Then

                Select Case potencia_total
                    Case <= 0
                        Exit Select
                    Case < 20000
                        secao = 10.0
                        disjuntor = 50
                    Case < 25000
                        secao = 16.0
                        disjuntor = 63
                End Select

            ElseIf esquema = "3F+N+T" Then

                Select Case potencia_total
                    Case <= 0
                        Exit Select
                    Case < 25000
                        secao = 6.0
                        disjuntor = 32
                    Case < 30000
                        secao = 10.0
                        disjuntor = 40
                    Case < 35000
                        secao = 10.0
                        disjuntor = 50
                    Case < 40000
                        secao = 16.0
                        disjuntor = 63
                    Case < 50000
                        secao = 25.0
                        disjuntor = 70
                    Case < 65000
                        secao = 35.0
                        disjuntor = 100
                    Case < 75000
                        secao = 50.0
                        disjuntor = 125
                End Select

            Else

                MsgBox("Esquema não definido")

            End If

        End If

    End Sub

    Private Sub OrdenaCircuitos()

        circuitos.Sort(Function(x, y) x.GetNumero().CompareTo(y.GetNumero())) 'Ordena lista

    End Sub

    Private Sub DistribuiFases()

        For i = 0 To (circuitos.Count - 1)   ' Runs through the circuits in the load table

            Dim circuito_atual As Circuito = circuitos(i)

            ' If it's the first circuit, don't check the previous circuits ('cause there ain't none)
            If i = 0 Then

                Select Case circuito_atual.GetEsquema()
                    Case "F+N+T", ""
                        circuito_atual.SetFases("R")
                        GoTo end_of_for_01
                    Case "2F+N+T"
                        circuito_atual.SetFases("R+S")
                        GoTo end_of_for_01
                    Case "3F+N+T"
                        circuito_atual.SetFases("R+S+T")
                        GoTo end_of_for_01
                End Select

            Else                                                    ' If else checks the previous circuits, unless it's a three phase

                If circuito_atual.GetEsquema().Contains("3F+N+T") Then

                    circuito_atual.SetFases("R+S+T")
                    GoTo end_of_for_01

                Else

                    For i_anterior = (i - 1) To 0 Step -1

                        Dim circuito_anterior As Circuito = circuitos(i_anterior)

                        If circuito_anterior.GetPotenciaR() > 0 And circuito_anterior.GetPotenciaS() > 0 And circuito_anterior.GetPotenciaT() > 0 Then

                            GoTo end_of_for_02

                        Else

                            If circuito_anterior.GetPotenciaR() > 0 Then

                                If circuito_anterior.GetPotenciaS() > 0 Then

                                    Select Case circuito_atual.GetEsquema()

                                        Case "F+N+T", ""
                                            circuito_atual.SetFases("T")
                                            GoTo end_of_for_01
                                        Case "2F+N+T"
                                            circuito_atual.SetFases("R+T")
                                            GoTo end_of_for_01
                                    End Select

                                Else

                                    Select Case circuito_atual.GetEsquema()
                                        Case "F+N+T", ""
                                            circuito_atual.SetFases("S")
                                            GoTo end_of_for_01
                                        Case "2F+N+T"
                                            circuito_atual.SetFases("S+T")
                                            GoTo end_of_for_01
                                    End Select

                                End If

                            ElseIf circuito_anterior.GetPotenciaS() > 0 Then

                                If circuito_anterior.GetPotenciaT > 0 Then

                                    Select Case circuito_atual.GetEsquema()
                                        Case "F+N+T", ""
                                            circuito_atual.SetFases("R")
                                            GoTo end_of_for_01
                                        Case "2F+N+T"
                                            circuito_atual.SetFases("R+S")
                                            GoTo end_of_for_01
                                    End Select

                                Else

                                    Select Case circuito_atual.GetEsquema()
                                        Case "F+N+T", ""
                                            circuito_atual.SetFases("T")
                                            GoTo end_of_for_01
                                        Case "2F+N+T"
                                            circuito_atual.SetFases("R+T")
                                            GoTo end_of_for_01
                                    End Select

                                End If

                            Else

                                Select Case circuito_atual.GetEsquema()
                                    Case "F+N+T", ""
                                        circuito_atual.SetFases("R")
                                        GoTo end_of_for_01
                                    Case "2F+N+T"
                                        circuito_atual.SetFases("R+S")
                                        GoTo end_of_for_01
                                End Select

                            End If

                        End If

end_of_for_02:

                    Next

                End If

            End If

end_of_for_01:

        Next

    End Sub

    <Obsolete>
    Private Function MontaQC(posicao As Point3d)

        Dim tabela As Table = New Table()

        With tabela

            Dim headerRows As Integer = 2 'Quantidade de linhas para o cabeçalho
            Dim totalRows As Integer = 1 'Quantidade de linhas para o total
            Dim rows As Integer = circuitos.Count + headerRows + totalRows 'Quantidade de linhas totais (cabeçalho + conteúdo + total)
            Dim columnTextList() As String = {"Circuito", "Descrição", "Potência (W)", "R (W)", "S (W)", "T (W)", "Seção (mm²)", "Disjuntor"} 'Conteúdo das colunas
            Dim columns As Integer = UBound(columnTextList) - LBound(columnTextList) + 1 'Quantidade de colunas totais (baseado no tamanho do array do conteúdo das colunas)

            .SetSize(rows, columns)
            .SetRowHeight(16)
            .SetColumnWidth(100)
            .Position = posicao

            .SetTextHeight(0, 0, 8)
            .SetAlignment(0, 0, CellAlignment.MiddleCenter)
            .SetTextString(0, 0, "Quadro de Cargas")

            For indexColumn = 0 To (columns - 1)

                .SetTextHeight(1, indexColumn, 8)
                .SetAlignment(1, indexColumn, CellAlignment.MiddleCenter)
                .SetTextString(1, indexColumn, columnTextList(indexColumn))

            Next

            For indexRow = headerRows To (rows - totalRows - 1)

                For indexColumn = 0 To (columns - 1)

                    .SetTextHeight(indexRow, indexColumn, 8)
                    .SetAlignment(indexRow, indexColumn, CellAlignment.MiddleCenter)

                    Dim circRow = indexRow - headerRows

                    Select Case indexColumn
                        Case 0
                            .SetTextString(indexRow, indexColumn, circuitos(circRow).GetNumero().ToString())
                        Case 1
                            .SetTextString(indexRow, indexColumn, "??")
                            .SetColumnWidth(indexColumn, 300)
                        Case 2
                            .SetTextString(indexRow, indexColumn, circuitos(circRow).GetPotenciaTotal().ToString())
                        Case 3
                            .SetTextString(indexRow, indexColumn, circuitos(circRow).GetPotenciaR().ToString())
                        Case 4
                            .SetTextString(indexRow, indexColumn, circuitos(circRow).GetPotenciaS().ToString())
                        Case 5
                            .SetTextString(indexRow, indexColumn, circuitos(circRow).GetPotenciaT().ToString())
                        Case 6
                            .SetTextString(indexRow, indexColumn, circuitos(circRow).GetSecao().ToString())
                        Case 7
                            .SetTextString(indexRow, indexColumn, circuitos(circRow).GetDisjuntor().ToString())
                    End Select

                Next

            Next

            'Preenche a última linha (total)
            For indexColumn = 0 To (columns - 1)
                .SetTextHeight(rows - totalRows, indexColumn, 8)
                .SetAlignment(rows - totalRows, indexColumn, CellAlignment.MiddleCenter)
            Next
            .SetTextString(rows - totalRows, 0, "Total")
            Dim range As CellRange = CellRange.Create(tabela, rows - totalRows, 0, rows - totalRows, 1)
            .MergeCells(range)
            .SetTextString(rows - totalRows, 2, potencia_total.ToString)
            .SetTextString(rows - totalRows, 3, potencia_r.ToString)
            .SetTextString(rows - totalRows, 4, potencia_s.ToString)
            .SetTextString(rows - totalRows, 5, potencia_t.ToString)
            .SetTextString(rows - totalRows, 6, secao.ToString)
            .SetTextString(rows - totalRows, 7, disjuntor.ToString)

            tabela.GenerateLayout()

        End With

        Return tabela

    End Function

    <Obsolete>
    Public Sub DesenhaQC(posicao As Point3d)

        'Acessa a base de dados
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        Dim tabela As New Table
        tabela = MontaQC(posicao)
        If tabela <> Nothing Then

            Using trans As Transaction = doc.TransactionManager.StartTransaction()

                Call Canvas.TrocaLayer("MD - Quadro de Cargas")

                Dim bt As BlockTable = CType(trans.GetObject(doc.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                btr.AppendEntity(tabela)
                trans.AddNewlyCreatedDBObject(tabela, True)

                trans.Commit()

            End Using

        Else
            MsgBox("Erro ao gerar a tabela: A tabela é nula")
        End If

    End Sub

    Public Sub DesenhaDU(posicao As Point3d)

        'Caso o número de circuitos seja pequeno, define 
        Dim posicaoInferior As Double
        If circuitos.Count < 5 Then
            posicaoInferior = 5
        Else
            posicaoInferior = circuitos.Count
        End If


        'Muda a layer atual para a layer do Diagrama Unifilar
        Call Canvas.TrocaLayer("MD - Diagrama Unifilar")

        'Desenha o barramento que conecta os circuitos
        DesenhaLinha(New Point3d(posicao.X, posicao.Y + 20, posicao.Z), New Point3d(posicao.X, (posicao.Y - (circuitos.Count - 1) * 60) - 20, posicao.Z), 6)

        'Desenha o circuito
        For i = 0 To (circuitos.Count - 1)

            circuitos(i).Desenha(New Point3d(posicao.X, posicao.Y + (i * (-60)), posicao.Z))

        Next

        'Desenha os textos do nome do QD e da potência total, respectivamente
        DesenhaTexto(nome, New Point3d((posicao.X - 200), (posicao.Y + 65), 0), 7, 15) 'Texto da Descrição
        DesenhaTexto("(" & potencia_total.ToString & " W)", New Point3d((posicao.X - 195), (posicao.Y + 45), 0), 7, 10) 'Texto da Potência Total

        'Desenha o retângulo que envolve o Diagrama Unifilar
        Dim frame As New List(Of Point2d)
        frame.Add(New Point2d(posicao.X - 200, posicao.Y + 60))
        frame.Add(New Point2d(posicao.X + 125, posicao.Y + 60))
        frame.Add(New Point2d(posicao.X + 125, posicao.Y - (60 * (posicaoInferior - 1) + 200)))
        frame.Add(New Point2d(posicao.X - 200, posicao.Y - (60 * (posicaoInferior - 1) + 200)))
        frame.Add(New Point2d(posicao.X - 200, posicao.Y + 60))
        DesenhaPolyline(frame, 0, 0, 0, True, "Continuous")

        'Desenha o detalhe do terra
        DesenhaTerra(New Point3d(posicao.X + 75, posicao.Y - (60 * (posicaoInferior - 1) + 200), 0)) 'Desenha o Terra do Diagrama Unifilar

        'Desenha o barramento
        barramento.DesenhaBarramento(New Point3d(posicao.X, posicao.Y - (60 * (posicaoInferior - 1)) - 70, posicao.Z))

        'Desenha a linha que conecta o barramento dos circuitos ao ramal de entrada (coluna onde ficariam os DRs)
        DesenhaLinha(New Point3d(posicao.X, (posicao.Y + (posicao.Y - (circuitos.Count - 1) * 60)) / 2, posicao.Z),
                     New Point3d(posicao.X - 55, (posicao.Y + (posicao.Y - (circuitos.Count - 1) * 60)) / 2, posicao.Z), 7)

        'Desenha o ramal de entrada
        DesenhaRamalEntrada(New Point3d(posicao.X - 55, (posicao.Y + (posicao.Y - (circuitos.Count - 1) * 60)) / 2, posicao.Z))

    End Sub

    Public Sub DesenhaTerra(posicao As Point3d)

        DesenhaLinha(posicao, New Point3d(posicao.X, posicao.Y - 20, 0), 3)
        DesenhaLinha(New Point3d(posicao.X - 15, posicao.Y - 20, 0), New Point3d(posicao.X + 15, posicao.Y - 20, 0), 3)
        DesenhaLinha(New Point3d(posicao.X - 9, posicao.Y - 27.5, 0), New Point3d(posicao.X + 9, posicao.Y - 27.5, 0), 3)
        DesenhaLinha(New Point3d(posicao.X - 3, posicao.Y - 35, 0), New Point3d(posicao.X + 3, posicao.Y - 35, 0), 3)

    End Sub

    Public Sub DesenhaRamalEntrada(posicao As Point3d)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Opens the Block table for read
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Dim blkID As ObjectId = ObjectId.Null

            Dim nomeBloco As String = "MD - DU Proteção Entrada"
            'Checa se o bloco do ramal de entrada existe no .dwg atual
            If Not bt.Has(nomeBloco) Then
                'Adiciona o bloco a partir de um arquivo .dwg
                MsgBox("Bloco " & nomeBloco & " não encontrado. Selecione o arquivo .dwg que contenha este bloco: ")
                If ImportaBloco(nomeBloco) = True Then
                    blkID = bt(nomeBloco)
                Else
                    MsgBox("Bloco do barramento não importado corretamente!")
                    Return
                End If
            Else
                blkID = bt("MD - DU Proteção Entrada")
            End If

            'Cria e insere a referência do novo bloco
            If blkID <> ObjectId.Null Then

                Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)
                    Using blkRef As New BlockReference(posicao, btr.Id)

                        Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        curBtr.AppendEntity(blkRef)
                        trans.AddNewlyCreatedDBObject(blkRef, True)

                        'Verifica se o block Block Table Record possui definições de atributos associados a ele
                        If btr.HasAttributeDefinitions Then

                            'Adiciona atributos a partir do Block Table Record
                            For Each objID As ObjectId In btr

                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim attDef As AttributeDefinition = dbObj

                                    If Not attDef.Constant Then

                                        Using attRef As New AttributeReference

                                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                            attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                            Select Case attRef.Tag
                                                Case "DISJUNTOR"
                                                    attRef.TextString = disjuntor & " A - C"
                                                    Exit Select
                                                Case "SEÇÃO"
                                                    attRef.TextString = "[3#" & secao & "(" & secao & ")" & secao & "] mm²"
                                                    Exit Select
                                                Case "NOME_ENTRADA"
                                                    attRef.TextString = derivaDe
                                                    Exit Select
                                            End Select

                                            ' Add DU block to the block table record and to the transaction
                                            blkRef.AttributeCollection.AppendAttribute(attRef)
                                            trans.AddNewlyCreatedDBObject(attRef, True)

                                        End Using
                                    End If
                                End If
                            Next
                        End If
                    End Using
                End Using
            End If

            trans.Commit()

        End Using

    End Sub

    'Public Sub PerguntaNomeQD()

    '    Dim msgAoUsuario As String = "Digite o nome do Quadro de Distribuição (ex. 'QD01 - Condomínio'): "
    '    PedeTextoUsuario(msgAoUsuario)

    'End Sub

End Class
