Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Windows
Imports System.IO
Imports System
Imports System.Windows

Public Module InterfaceACAD

    Public Function SelecionaCoordenada()

        'Acessa a base de dados
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Dim posicao As PromptPointResult = ed.GetPoint("Selecione o local onde a tabela será inserida: ")  ' Asks the user for the table's insertion point

        If posicao.Status = PromptStatus.OK Then
            Return posicao.Value
        Else
            MsgBox("Erro ao selecionar a posição. Tabela desenhada no ponto (0,0,0).")
            Return New Point3d(0, 0, 0)
        End If

    End Function

    Public Function RetornaPerimetro(ed As Editor)

        Dim peo As PromptEntityOptions = New PromptEntityOptions("Selecione a POLYLINE que delimita o perímetro dos circuitos do QD:")

        peo.SetRejectMessage("Objeto selecionado não é do tipo POLYLINE")
        peo.AddAllowedClass(GetType(Polyline), True)

        Dim perimetro = ed.GetEntity(peo)

        Return perimetro

    End Function

    Public Function RetornaObjetosDentroDePolyline(ed As Editor, trans As Transaction, perimetro As PromptEntityResult)

        Dim pl As Polyline = CType(trans.GetObject(perimetro.ObjectId, OpenMode.ForRead), Polyline)
        Dim vertices(pl.NumberOfVertices) As Point3d 'Gets the vertices of the perimeter
        Dim pointCollection As Point3dCollection = New Point3dCollection()

        'Passes the vertices of the perimeter to a point3dCollection
        For i = 0 To pl.NumberOfVertices - 1

            pointCollection.Add(pl.GetPoint3dAt(i))

        Next

        Dim tv As TypedValue() = New TypedValue(0) {}
        tv.SetValue(New TypedValue(CInt(DxfCode.Start), "INSERT"), 0)
        Dim filter As SelectionFilter = New SelectionFilter(tv)

        Dim ss As SelectionSet = ed.SelectCrossingPolygon(pointCollection, filter).Value 'Creates a selection using the perimeter as reference

        Return ss

    End Function

    Public Function PedeTextoUsuario(msgAoUsuario As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        Dim pStrOpts As PromptStringOptions = New PromptStringOptions(msgAoUsuario)
        pStrOpts.AllowSpaces = True
        Dim inputUsuario As PromptResult = doc.Editor.GetString(pStrOpts)

        Return inputUsuario.StringResult

    End Function

    'Retorna o index do array opcoes
    Public Function PedeInputUsuario(opcoes() As String)

        'Dim opcoes() As String = {"Procurar bloco em outro arquivo.", "Cancelar procura (Diagrama Unifilar não será representado corretamente)."}

        If opcoes.Count > 0 Then

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument

            Dim msgText As String = ""
            For i = 0 To opcoes.Count - 1
                Dim opcao As Integer = i + 1
                msgText = msgText & vbNewLine & opcao & " - " & opcoes(i)
            Next

            Dim pStrOpts As PromptStringOptions = New PromptStringOptions(msgText)
            pStrOpts.AllowSpaces = True
            Dim inputUsuario As PromptResult = doc.Editor.GetString(pStrOpts)

            Dim inputValido As Boolean = False
            For i = 0 To opcoes.Count - 1
                If inputUsuario.StringResult = (i + 1).ToString Then
                    inputValido = True
                End If
            Next

            If inputValido = True Then
                'MsgBox(inputUsuario.StringResult)
                Return inputUsuario.StringResult
            Else
                MsgBox("Input inválido!")
                PedeInputUsuario(opcoes)
            End If

            'Application.ShowAlertDialog("The name entered was: " & pStrRes.StringResult)

        Else
            MsgBox("Nenhuma opção adicionada!")
        End If

        Return Nothing

    End Function

    <CommandMethod("ImportaBloco")>
    Public Function ImportaBloco(nomeBloco As String)

        Dim name As String = Application.DocumentManager.MdiActiveDocument.Name

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        Using OpenDb As New Database(False, True)

            Dim ofd As New OpenFileDialog("Selecione o Arquivo .dwg que contenha o bloco " & nomeBloco & ": ", "", "dwg", "Importador de Blocos",
                                      OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles) 'Configura o Form do Explorador de Arquivos
            Dim dialogResult As System.Windows.Forms.DialogResult = ofd.ShowDialog() 'Abre o Explorador de Arquivos
            Dim filePath As String = ofd.Filename
            OpenDb.ReadDwgFile(filePath, System.IO.FileShare.ReadWrite, True, "")

            Dim ids As ObjectIdCollection = New ObjectIdCollection()

            Using tr As Transaction = OpenDb.TransactionManager.StartTransaction()

                Dim bt As BlockTable = tr.GetObject(OpenDb.BlockTableId, OpenMode.ForRead)

                If (bt.Has(nomeBloco)) Then
                    MsgBox("Bloco encontrado!")
                    ids.Add(bt(nomeBloco))
                    tr.Commit()
                Else
                    MsgBox("Bloco não encontrado!")
                    Dim opcoes() As String = {"Procurar bloco em outro arquivo.", "Cancelar procura (Diagrama Unifilar não será representado corretamente)."}
                    Dim inputUsuario As String = PedeInputUsuario(opcoes)
                    If inputUsuario = 1 Then
                        ImportaBloco(nomeBloco)
                    Else
                        Return False
                    End If
                End If

            End Using


            'Adiciona o bloco se encontrado

            If (ids.Count <> 0) Then

                MsgBox("Adicionando Bloco.")

                'Acessa a base de dados

                Dim destdb As Database = doc.Database

                Dim iMap As IdMapping = New IdMapping()

                destdb.WblockCloneObjects(ids, destdb.BlockTableId, iMap, DuplicateRecordCloning.Replace, False)

                Return True

            End If

            Return False

        End Using

    End Function

End Module
