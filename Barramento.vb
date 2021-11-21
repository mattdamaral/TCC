Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class Barramento

    Private disjuntor As Integer
    Private tamanho As String

    Public Sub New(_disjuntor As Integer)

        disjuntor = _disjuntor
        DimensionaBarramento()

    End Sub

    Public Sub DimensionaBarramento()

        If disjuntor <> Nothing Then

            Select Case (disjuntor * 1.3)
                Case < 140
                    tamanho = "15x2"
                    Exit Select
                Case < 170
                    tamanho = "15x3"
                    Exit Select
                Case < 185
                    tamanho = "20x2"
                    Exit Select
                Case < 220
                    tamanho = "20x3"
                    Exit Select
                Case < 270
                    tamanho = "25x3"
                    Exit Select
                Case < 295
                    tamanho = "20x5"
                    Exit Select
                Case < 315
                    tamanho = "30x3"
                    Exit Select
                Case < 350
                    tamanho = "25x5"
                    Exit Select
                Case < 400
                    tamanho = "30x5"
                    Exit Select
                Case < 420
                    tamanho = "40x3"
                    Exit Select
                Case < 520
                    tamanho = "40x5"
                    Exit Select
                Case < 630
                    tamanho = "50x5"
                    Exit Select
                Case < 760
                    tamanho = "40x10"
                    Exit Select
                Case < 820
                    tamanho = "50x10"
                    Exit Select
                Case < 970
                    tamanho = "80x5"
                    Exit Select
                Case < 1060
                    tamanho = "60x10"
                    Exit Select
                Case < 1200
                    tamanho = "100x5"
                    Exit Select
                Case < 1380
                    tamanho = "80x10"
                    Exit Select
                Case < 1700
                    tamanho = "100x10"
                    Exit Select
                Case < 2000
                    tamanho = "120x10"
                    Exit Select
                Case < 2500
                    tamanho = "160x10"
                    Exit Select
                Case < 3000
                    tamanho = "200x10"
                    Exit Select

            End Select

        End If

    End Sub

    ' Draws the busbar (dimensions for the phase/neutral/ground bars)
    Public Sub DesenhaBarramento(posicao As Point3d)

        'Dim corFiacao As Integer = 7 'White

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        ' Starts a transaction
        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Opens the Block table for read
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Dim blkID As ObjectId = ObjectId.Null ' The ID of the DR block

            Dim nomeBloco As String = "MD - DU Barramento"
            'Checa se o bloco do barramento existe no .dwg atual
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
                blkID = bt(nomeBloco)
            End If

            ' Creates and inserts the new block reference
            If blkID <> ObjectId.Null Then

                Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)

                    Using blkRef As New BlockReference(posicao, btr.Id)

                        Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        curBtr.AppendEntity(blkRef)
                        trans.AddNewlyCreatedDBObject(blkRef, True)

                        ' Verify block table record has attribute definitions associated with it
                        If btr.HasAttributeDefinitions Then
                            For Each objID As ObjectId In btr

                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim attDef As AttributeDefinition = dbObj

                                    If Not attDef.Constant Then

                                        Using attRef As New AttributeReference

                                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                            attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                            'Checks If the attribute's tag is one of the below
                                            Select Case attRef.Tag
                                                Case "F"
                                                    attRef.TextString = "F - " + tamanho + " mm"
                                                Case "N"
                                                    attRef.TextString = "N - " + tamanho + " mm"
                                                Case "PE"
                                                    attRef.TextString = "PE - " + tamanho + " mm"
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

    Public Function GetTamanho()
        Return tamanho
    End Function

End Class
