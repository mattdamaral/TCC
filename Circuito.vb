Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class Circuito

    Private numero As String
    Private esquema As String
    'Private tensao As String
    Private potencia_total As Double
    Private fases As String
    Private potencia_r As Double
    Private potencia_s As Double
    Private potencia_t As Double
    Private secao As Double
    Private disjuntor As Integer
    Private conexao As Integer

    Public Sub New(_numero As String, _conexao As Integer, _potencia_total As Double, _secao As Double, _disjuntor As Integer)

        numero = _numero
        conexao = _conexao
        potencia_total = _potencia_total
        secao = _secao
        disjuntor = _disjuntor

        potencia_r = 0
        potencia_s = 0
        potencia_t = 0

        DefineEsquema()

    End Sub

    Private Sub DefineEsquema()

        If conexao = "1" Then
            esquema = "F+N+T"
        ElseIf conexao = "2" Then
            esquema = "2F+N+T"
        ElseIf conexao = "3" Then
            esquema = "3F+N+T"
        Else
            MsgBox("Esquema do circuito " + numero.ToString + "não definido corretamente.")
            esquema = "??"
        End If

    End Sub

    Public Sub SetFases(_fases As String)
        fases = _fases
        DividePotenciaFases()
    End Sub

    Private Sub DividePotenciaFases()

        Dim fasesArray As String() = fases.Split(CChar("+"))

        If fasesArray.Count = conexao Then
            For i = 0 To (fasesArray.Count - 1)
                Select Case fasesArray(i)
                    Case "R"
                        potencia_r = potencia_total / fasesArray.Count
                        Exit Select
                    Case "S"
                        potencia_s = potencia_total / fasesArray.Count
                        Exit Select
                    Case "T"
                        potencia_t = potencia_total / fasesArray.Count
                        Exit Select
                End Select
            Next
        Else
            MsgBox("Quantidade de fases do circuito " + numero.ToString + "não condiz com o esquema de ligação.")
        End If

    End Sub

    Public Sub Desenha(posicao As Point3d)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            'Abre o Block Table para leitura 
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim blkID As ObjectId = ObjectId.Null
            Dim nomeBloco As String = "MD - DU Circuito"

            'Checa se o bloco do circuito faz parte do projeto ou não
            If Not bt.Has(nomeBloco) Then 'Se o Bloco não faz parte do projeto, adiciona ele
                'Adicionar Bloco pelo .dwg caso já não faça parte do projeto
                MsgBox("Bloco " & nomeBloco & "não encontrado. O Diagrama Unifilar não será representado corretamente.")
                ImportaBloco(nomeBloco)
                blkID = bt(nomeBloco)
            Else 'Se o Bloco faz parte do projeto, utiliza ele
                blkID = bt(nomeBloco)
            End If

            If blkID <> ObjectId.Null Then

                Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead) 'Block Table Record do Bloco do Circuito

                    Using blkRef As New BlockReference(posicao, btr.Id) 'Cria o Block Reference do Bloco do Circuito - Leva em consideração sua posição de inserção

                        Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        curBtr.AppendEntity(blkRef) 'Insere o Block Reference ao Model Space
                        trans.AddNewlyCreatedDBObject(blkRef, True) 'Confirma a inserção

                        'Verifica se o Bloco possui Atributos
                        If btr.HasAttributeDefinitions Then

                            'Checa Atributo por Atributo
                            For Each objID As ObjectId In btr

                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim attDef As AttributeDefinition = dbObj

                                    If Not attDef.Constant Then

                                        Using attRef As New AttributeReference

                                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                            attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                            Select Case attRef.Tag
                                                Case "CIRC_N_CIRCUITO"
                                                    attRef.TextString = numero.ToString
                                                    Exit Select
                                                Case "CIRC_DESCRIÇÃO"
                                                    attRef.TextString = "(??)"
                                                    Exit Select
                                                Case "CIRC_POTÊNCIA"
                                                    attRef.TextString = "(" + potencia_total.ToString + " W)"
                                                    Exit Select
                                                Case "CIRC_FASE"
                                                    attRef.TextString = fases
                                                    Exit Select
                                                Case "CIRC_SEÇÃO"
                                                    attRef.TextString = secao.ToString & " mm²"
                                                    Exit Select
                                                Case "CIRC_DISJUNTOR"
                                                    'If desc.Contains("Iluminação") Then
                                                    '    attRef.TextString = disj.ToString & " A - B"
                                                    'Else
                                                    '    attRef.TextString = disj.ToString & " A - C"
                                                    'End If
                                                    attRef.TextString = disjuntor.ToString & " A"
                                                    Exit Select
                                            End Select

                                            blkRef.AttributeCollection.AppendAttribute(attRef) 'Adiciona o Bloco do Circuito ao Canvas
                                            trans.AddNewlyCreatedDBObject(attRef, True) 'Confirma a Transação

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

    Public Function GetNumero()
        Return numero
    End Function

    Public Function GetEsquema()
        Return esquema
    End Function

    Public Function GetPotenciaTotal()
        Return potencia_total
    End Function

    Public Sub SetPotenciaTotal(_potencia_total As Double)
        potencia_total = _potencia_total
    End Sub

    Public Sub AdicionaAPotenciaTotal(_potencia As Double)
        potencia_total += _potencia
    End Sub

    Public Function GetPotenciaR()
        Return potencia_r
    End Function

    Public Sub SetPotenciaR(_potencia_r As Double)
        potencia_r = _potencia_r
    End Sub

    Public Function GetPotenciaS()
        Return potencia_s
    End Function

    Public Sub SetPotenciaS(_potencia_s As Double)
        potencia_s = _potencia_s
    End Sub

    Public Function GetPotenciaT()
        Return potencia_t
    End Function

    Public Sub SetPotenciaT(_potencia_T As Double)
        potencia_t = _potencia_T
    End Sub

    Public Function GetSecao()
        Return secao
    End Function

    Public Sub SetSecao(_secao As Double)
        secao = _secao
    End Sub

    Public Function GetDisjuntor()
        Return disjuntor
    End Function

    Public Sub SetDisjuntor(_disjuntor As Integer)
        disjuntor = _disjuntor
    End Sub

End Class
