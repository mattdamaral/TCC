Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Imports System.IO

Public Module Main

    <CommandMethod("DesenhaQCDU")>
    Public Sub DesenhaQCDU()

        Dim dados_suficientes As Boolean = False

        'Acessa a base de dados do AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = doc.TransactionManager.StartTransaction()

            'Tenta selecionar a Polyline que delimita o perímetro do projeto

            Dim perimetro As PromptEntityResult = RetornaPerimetro(ed)

            If perimetro.Status = PromptStatus.OK Then

                Dim objetos As SelectionSet = RetornaObjetosDentroDePolyline(ed, trans, perimetro)

                If objetos IsNot Nothing Then

                    Dim blocos As New List(Of BlockReference)

                    For Each objeto As SelectedObject In objetos

                        Dim bloco As BlockReference = CType(trans.GetObject(objeto.ObjectId, OpenMode.ForRead), BlockReference)
                        blocos.Add(bloco)

                    Next

                    ''Pergunta o nome do Quadro de Distribuição
                    'Dim msgAoUsuario As String = "Digite o nome do Quadro de Distribuição (ex. 'QD01 - Condomínio'): "
                    'PedeTextoUsuario(msgAoUsuario)

                    Dim qd As New QD(PedeTextoUsuario("Digite o nome do Quadro de Distribuição (ex.: 'QD01 - Condomínio'): "),
                                     PedeTextoUsuario("Digite o nome do Quadro de ondem este QD deriva (ex.: 'QM01'"),
                                     CriaCircuitos(blocos, trans), DefineMateriais(blocos, trans))

                    qd.DesenhaQC(SelecionaCoordenada("Selecione o local onde o Quadro de Cargas será inserido: "))
                    qd.DesenhaDU(SelecionaCoordenada("Selecione o local onde o Diagrama Unifilar será inserida: "))
                    qd.DesenhaLM(SelecionaCoordenada("Selecione o local onde a Lista de Materiais será inserida: "))

                    trans.Commit()

                Else

                    MsgBox("Erro na seleção da POLYLINE")

                End If

            End If

        End Using

    End Sub

    Private Function CriaCircuitos(blocos As List(Of BlockReference), trans As Transaction)

        Dim circuitos As New List(Of Circuito)

        For Each bloco As BlockReference In blocos

            Dim ac As AttributeCollection = bloco.AttributeCollection
            Dim numeroDeAtributosCompativeis As Integer = 0

            'Parâmetros de criação de um circuito
            Dim numero As String = ""
            Dim potencia As Double = 0
            Dim conexão As Integer = 1
            Dim secao As Double = 0
            Dim disjuntor As Integer = 0

            For Each objID As ObjectId In ac

                Dim atributo As AttributeReference = CType(trans.GetObject(objID, OpenMode.ForRead), AttributeReference)  'Armazena os atributos

                Select Case atributo.Tag
                    Case "CIRCUITO"
                        numeroDeAtributosCompativeis += 1
                        numero = atributo.TextString
                        Exit Select
                    Case "POTÊNCIA"
                        numeroDeAtributosCompativeis += 1
                        Double.TryParse(atributo.TextString, potencia)
                        Exit Select
                    Case "CONEXÃO"
                        numeroDeAtributosCompativeis += 1
                        Integer.TryParse(atributo.TextString, conexão)
                        Exit Select
                    Case "SEÇÃO"
                        numeroDeAtributosCompativeis += 1
                        Double.TryParse(atributo.TextString, secao)
                        Exit Select
                    Case "DISJUNTOR"
                        numeroDeAtributosCompativeis += 1
                        Integer.TryParse(atributo.TextString, disjuntor)
                        Exit Select
                End Select

            Next

            If numeroDeAtributosCompativeis = 5 Then

                If circuitos.Count > 0 Then 'Checa se já existe circuitos na lista de circuitos

                    Dim circuitoJaExiste = False

                    For i = 0 To (circuitos.Count - 1) 'Checa se o número do bloco é igual a um dos existentes

                        If numero = circuitos(i).GetNumero() Then 'Se já existe circuitos na lista de circuitos, e o número do bloco é igual a um dos existentes, adiciona/modifica os parâmetros

                            circuitoJaExiste = True
                            circuitos(i).AdicionaAPotenciaTotal(potencia)

                            If circuitos(i).GetSecao < secao Then
                                circuitos(i).SetSecao(secao)
                            End If

                            If circuitos(i).GetDisjuntor < disjuntor Then
                                circuitos(i).SetDisjuntor(disjuntor)
                            End If

                            GoTo goto_01

                        End If

                    Next

                    If circuitoJaExiste = False Then 'Se já existe circuitos na lista de circuitos, porém o número do bloco difere dos existentes, cria um novo

                        circuitos.Add(New Circuito(numero, conexão, potencia, secao, disjuntor))

                    End If

                Else 'Se a lista de circuitos está vazia, cria um novo

                    circuitos.Add(New Circuito(numero, conexão, potencia, secao, disjuntor))

                End If

            Else


            End If

goto_01:

        Next

        Return circuitos

    End Function

    Private Function DefineMateriais(blocos As List(Of BlockReference), trans As Transaction)

        Dim materiais As New List(Of Material)

        For Each bloco As BlockReference In blocos

            Dim ac As AttributeCollection = bloco.AttributeCollection
            Dim numeroDeAtributosCompativeis As Integer = 0

            'Parâmetros de criação de um material
            Dim nome As String
            Dim quantidade As Integer

            For Each objID As ObjectId In ac

                Dim atributo As AttributeReference = CType(trans.GetObject(objID, OpenMode.ForRead), AttributeReference)  'Armazena os atributos

                If atributo.Tag.Contains("LM_") Then
                    Dim nomeAux As String = atributo.Tag.Replace("LM_", "")
                    nome = nomeAux.Replace("_", " ")
                    Dim quantidadeDouble As Double
                    Double.TryParse(atributo.TextString(), quantidadeDouble)
                    Integer.TryParse(quantidadeDouble, quantidade)

                    If quantidade > 0 Then

                        If materiais.Count > 0 Then

                            For i = 0 To materiais.Count - 1

                                If materiais(i).GetNome() = nome Then

                                    materiais(i).AdicionaQuantidade(quantidade)

                                    GoTo goto_01

                                End If

                            Next

                            materiais.Add(New Material(nome, quantidade))

                        Else

                            materiais.Add(New Material(nome, quantidade))

                        End If

                        Exit For

                    End If

                End If

goto_01:

            Next

        Next

        Return materiais

    End Function

End Module
