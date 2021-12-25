Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Module Canvas

    'Desenha uma Linha entre 2 pontos 3D
    Public Sub DesenhaLinha(posInicial As Point3d, posFinal As Point3d, cor As Integer)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        'Inicia uma Transação
        Using trans As Transaction = db.TransactionManager.StartTransaction()

            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) 'Abre o Block Table para Leitura
            Using btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite) 'Abre o Block Table Record para Escrita

                Dim linha As New Line(posInicial, posFinal) 'Cria a Linha
                linha.ColorIndex = cor 'Define a Cor da Linha

                btr.AppendEntity(linha) 'Adiciona a Linha ao Canvas
                trans.AddNewlyCreatedDBObject(linha, True) 'Confirma a Adição da Linha

            End Using

            trans.Commit() 'Confirma a Transação

        End Using
    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Seleciona a Layer Atual
    Public Sub TrocaLayer(nomeLayer As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        'Inicia uma Transação
        Using trans As Transaction = doc.TransactionManager.StartTransaction()

            Dim lt As LayerTable
            Dim ltr As New LayerTableRecord
            Dim layerID As ObjectId

            'Checa se a Layer existe
            Try

                'Se a Layer existe, obtém seu ID
                lt = CType(trans.GetObject(db.LayerTableId, OpenMode.ForRead, True, True), LayerTable)
                layerID = lt.Item(nomeLayer)

                'Se a Layer foi deletada, recupera a Layer
                If layerID.IsErased Then
                    lt.UpgradeOpen()
                    lt.Item(nomeLayer).GetObject(OpenMode.ForWrite, True, True).Erase(False)
                End If

            Catch ex As Autodesk.AutoCAD.Runtime.Exception

                'Se a Layer não existe, cria uma nova
                lt = db.LayerTableId.GetObject(OpenMode.ForWrite, True, True)
                ltr.Name = nomeLayer
                lt.Add(ltr)
                'Adiciona a Layer ao Database
                trans.AddNewlyCreatedDBObject(ltr, True)
                'Obtém o ID da Layer criada
                lt = CType(trans.GetObject(db.LayerTableId, OpenMode.ForRead, False), LayerTable)
                layerID = lt.Item(nomeLayer)

            End Try

            'Define a Layer como a Layer Atual
            db.Clayer = layerID

            trans.Commit()

        End Using

    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub DesenhaTexto(texto As String, posicao As Point3d, cor As Integer, altura As Double)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Using btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                Dim dbText As DBText = New DBText
                dbText.SetDatabaseDefaults()
                dbText.TextString = texto
                dbText.ColorIndex = cor
                dbText.Position = posicao
                dbText.Height = altura
                btr.AppendEntity(dbText)
                trans.AddNewlyCreatedDBObject(dbText, True)

            End Using

            trans.Commit() 'Confirma a Transação

        End Using

    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub DesenhaPolyline(posVertices As List(Of Point2d), bulge As Double, startWidth As Double, endWidth As Double, isClosed As Boolean, lineType As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Using btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                Dim pl As Polyline = New Polyline

                For i = 0 To posVertices.Count - 1

                    pl.AddVertexAt(i, posVertices(i), bulge, startWidth, endWidth)

                Next

                pl.Closed = isClosed
                pl.Linetype = lineType

                btr.AppendEntity(pl)
                trans.AddNewlyCreatedDBObject(pl, True)

            End Using

            trans.Commit() 'Confirma a Transação

        End Using
    End Sub

    Public Sub DesenhaMText(texto As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Using btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                'Dim mleader As New MLeader
                'mleader.SetDatabaseDefaults()
                'mleader.ContentType = ContentType.MTextContent
                'Dim mtext As New MText
                'mtext.SetDatabaseDefaults()
                'mtext.Height = 8
                'mtext.SetContentsRtf(texto)
                ''... https://adndevblog.typepad.com/autocad/2012/05/how-to-create-mleader-objects-in-net.html


                Dim mtext As New MText
                mtext.SetDatabaseDefaults()
                mtext.Contents = texto
                mtext.ColorIndex = 7
                mtext.Location = SelecionaCoordenada("Selecione o local onde será inserido o texto com a sugestão: ")
                mtext.TextHeight = 8
                btr.AppendEntity(mtext)
                trans.AddNewlyCreatedDBObject(mtext, True)

                'Dim mtext As New MText
                'mtext.SetDatabaseDefaults()
                'mtext.BackgroundFill = True
                'mtext.UseBackgroundColor = True
                'mtext.TextHeight = 8
                'mtext.Contents = texto

                'Dim posicao As New Point3d
                'posicao = SelecionaCoordenada()

                'mtext.Location = posicao

                'btr.AppendEntity(mtext)
                'trans.AddNewlyCreatedDBObject(mtext, True)

                trans.Commit()

            End Using

        End Using

    End Sub

    <CommandMethod("MedePolyline")>
    Public Sub MedePolyline()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            Dim pl As Polyline = CType(trans.GetObject(RetornaPerimetro(ed).ObjectId, OpenMode.ForRead), Polyline)

            Dim area As Double = (pl.Area) / 10000 'Área englobada pela polyline em metros²
            Dim perimetro As Double = (pl.Length) / 100 'Perímetro da polyline em metros

            DesenhaMText("Área: " + area.ToString + " m²" + vbNewLine + "Perímetro: " + perimetro.ToString + " m")

            trans.Commit()

        End Using

    End Sub

End Module
