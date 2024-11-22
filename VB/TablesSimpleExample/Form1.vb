Imports DevExpress.Office.Utils
Imports DevExpress.XtraBars
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Drawing

Namespace TablesSimpleExample

    Public Partial Class Form1
        Inherits RibbonForm

        Private table As Table

        Private document As Document

        Public Sub New()
            InitializeComponent()
            document = richEditControl1.Document
        End Sub

        Private Sub createTablebtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            ' Create a new table and specify its layout type
            table = document.Tables.Create(document.Range.End, 2, 2)
            ' Add new rows to the table
            table.Rows.InsertBefore(0)
            table.Rows.InsertAfter(0)
            ' Add a new column to the table
            table.Rows(0).Cells.Append()
            table.Rows(0).Cells.InsertBefore(0)
            table.Rows(0).FirstCell.PreferredWidthType = WidthType.Auto
            table.Rows(0).Cells(1).PreferredWidthType = WidthType.Fixed
            table.Rows(0).Cells(1).PreferredWidth = Units.InchesToDocumentsF(0.8F)
            ' Set the second column width and cell height
            table(0, 2).PreferredWidthType = WidthType.Fixed
            table(0, 2).PreferredWidth = Units.InchesToDocumentsF(5F)
            table(0, 2).HeightType = HeightType.Exact
            table(0, 2).Height = Units.InchesToDocumentsF(0.5F)
            'Set the third column width 
            table.Rows(0).LastCell.PreferredWidthType = WidthType.Fixed
            table.Rows(0).LastCell.PreferredWidth = Units.InchesToDocumentsF(0.8F)
        End Sub

        Private Sub mergeBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            table.BeginUpdate()
            ' Split cell into 8
            table(3, 2).Split(4, 2)
            ' Merge cells
            table.MergeCells(table(4, 2), table(4, 3))
            table.MergeCells(table(6, 2), table(6, 3))
            table.MergeCells(table(2, 0), table(6, 0))
            table.EndUpdate()
        End Sub

        Private Sub repeatRowsBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            table.BeginUpdate()
            ' Repeat first three rows as header:
            table.Rows(0).RepeatAsHeaderRow = True
            table.Rows(1).RepeatAsHeaderRow = True
            table.Rows(2).RepeatAsHeaderRow = True
            ' Break last row across pages:
            table.LastRow.BreakAcrossPages = True
            table.EndUpdate()
        End Sub

        Private Sub wrapTextBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            table.BeginUpdate()
            ' Wrap text around the table
            table.TextWrappingType = TableTextWrappingType.Around
            ' Specify vertical alignment:
            table.RelativeVerticalPosition = TableRelativeVerticalPosition.Paragraph
            table.VerticalAlignment = TableVerticalAlignment.None
            table.OffsetYRelative = Units.InchesToDocumentsF(2F)
            ' Specify horizontal alignment:
            table.RelativeHorizontalPosition = TableRelativeHorizontalPosition.Margin
            table.HorizontalAlignment = TableHorizontalAlignment.Center
            ' Set distance between the text and the table:
            table.MarginBottom = Units.InchesToDocumentsF(0.3F)
            table.MarginLeft = Units.InchesToDocumentsF(0.3F)
            table.MarginTop = Units.InchesToDocumentsF(0.3F)
            table.MarginRight = Units.InchesToDocumentsF(0.3F)
            table.EndUpdate()
        End Sub

        Private Sub insertContentBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            ' Insert header data
            document.InsertSingleLineText(table.Rows(0).Cells(2).Range.Start, "Active Customers")
            document.InsertSingleLineText(table(2, 1).Range.Start, "Photo")
            document.InsertSingleLineText(table(2, 0).Range.Start, "Customer №1")
            document.InsertSingleLineText(table(2, 2).Range.Start, "Customer Info")
            document.InsertSingleLineText(table(2, 3).Range.Start, "Rentals")
            ' Insert the customer photo
            document.Images.Insert(table(3, 1).Range.Start, DocumentImageSource.FromFile("photo.png"))
            ' Insert the customer info
            document.InsertText(table(3, 2).Range.Start, "Ryan Anita W")
            document.InsertText(table(3, 3).Range.Start, "Intermediate")
            document.InsertText(table(4, 2).Range.Start, "3/28/1984")
            document.InsertText(table(5, 2).Range.Start, "anita_ryan@dxvideorent.com")
            document.InsertText(table(5, 3).Range.Start, "(555)421-0059")
            document.InsertText(table(6, 2).Range.Start, "5119 Beryl Dr, San Antonio, TX 78212")
            document.InsertSingleLineText(table(3, 4).Range.Start, "18")
        End Sub

        Private Sub formatContentBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            ' Apply formatting to the "Active Customers" cell
            Dim properties As CharacterProperties = document.BeginUpdateCharacters(table(0, 2).ContentRange)
            properties.FontName = "Segoe UI"
            properties.FontSize = 16
            document.EndUpdateCharacters(properties)
            Dim alignment As ParagraphProperties = document.BeginUpdateParagraphs(table(0, 2).ContentRange)
            alignment.Alignment = ParagraphAlignment.Center
            document.EndUpdateParagraphs(alignment)
            table(0, 1).VerticalAlignment = TableCellVerticalAlignment.Center
            ' Apply formatting to the header cells
            Dim headerRowProperties As CharacterProperties = document.BeginUpdateCharacters(table.Rows(2).Range)
            headerRowProperties.FontName = "Segoe UI"
            headerRowProperties.FontSize = 11
            headerRowProperties.ForeColor = Color.FromArgb(212, 236, 183)
            document.EndUpdateCharacters(headerRowProperties)
            Dim headerRowParagraphProperties As ParagraphProperties = document.BeginUpdateParagraphs(table.Rows(2).Range)
            headerRowParagraphProperties.Alignment = ParagraphAlignment.Center
            document.EndUpdateParagraphs(headerRowParagraphProperties)
            ' Apply formatting to the customer info cells
            Dim targetRange As DocumentRange = document.CreateRange(table(3, 2).Range.Start, table(6, 3).Range.Start.ToInt() - table(3, 2).Range.Start.ToInt())
            Dim infoProperties As CharacterProperties = document.BeginUpdateCharacters(targetRange)
            infoProperties.FontSize = 8
            infoProperties.FontName = "Segoe UI"
            infoProperties.ForeColor = Color.FromArgb(111, 116, 106)
            document.EndUpdateCharacters(infoProperties)
            ' Format "Rentals" cells
            Dim rentalFormat As CharacterProperties = document.BeginUpdateCharacters(table(3, 4).Range)
            rentalFormat.FontSize = 28
            rentalFormat.Bold = True
            document.EndUpdateCharacters(rentalFormat)
            Dim rentalAlignment As ParagraphProperties = document.BeginUpdateParagraphs(table(3, 4).Range)
            rentalAlignment.Alignment = ParagraphAlignment.Center
            document.EndUpdateParagraphs(rentalAlignment)
            table(3, 4).VerticalAlignment = TableCellVerticalAlignment.Center
        End Sub

        Private Sub customizeBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            table.BeginUpdate()
            ' Call the ChangeCellBorderColor method
            ' for every cell in the first two rows
            For i As Integer = 0 To 2 - 1
                For j As Integer = 0 To table.Rows(i).Cells.Count - 1
                    ' Specify the border style and the background color for the header cells 
                    table(i, j).Borders.Bottom.LineStyle = BorderLineStyle.None
                    table(i, j).Borders.Left.LineStyle = BorderLineStyle.None
                    table(i, j).Borders.Right.LineStyle = BorderLineStyle.None
                    table(i, j).Borders.Top.LineStyle = BorderLineStyle.None
                    table(i, j).BackgroundColor = Color.Transparent
                Next
            Next

            Dim targetRow As TableRow = table.Rows(2)
            targetRow.Cells(1).BackgroundColor = Color.FromArgb(99, 122, 110)
            targetRow.Cells(2).BackgroundColor = Color.FromArgb(99, 122, 110)
            targetRow.Cells(3).BackgroundColor = Color.FromArgb(99, 122, 110)
            table.EndUpdate()
        End Sub

        Private Sub tableStyleBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            document.BeginUpdate()
            ' Create a new table style
            Dim tStyleMain As TableStyle = document.TableStyles.CreateNew()
            ' Specify style options
            Dim insideHorizontalBorder As TableBorder = tStyleMain.TableBorders.InsideHorizontalBorder
            insideHorizontalBorder.LineStyle = BorderLineStyle.Single
            insideHorizontalBorder.LineColor = Color.White
            Dim insideVerticalBorder As TableBorder = tStyleMain.TableBorders.InsideVerticalBorder
            insideVerticalBorder.LineStyle = BorderLineStyle.Single
            insideVerticalBorder.LineColor = Color.White
            tStyleMain.CellBackgroundColor = Color.FromArgb(227, 238, 220)
            tStyleMain.Name = "MyTableStyle"
            ' Add the style to the document collection
            document.TableStyles.Add(tStyleMain)
            ' Create conditional styles (styles for specific table elements)         
            Dim myNewStyleForOddRows As TableConditionalStyle = tStyleMain.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddRowBanding)
            myNewStyleForOddRows.CellBackgroundColor = Color.FromArgb(196, 220, 182)
            Dim myNewStyleForBottomRightCell As TableConditionalStyle = tStyleMain.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.BottomRightCell)
            myNewStyleForBottomRightCell.CellBackgroundColor = Color.FromArgb(188, 214, 201)
            document.EndUpdate()
            document.BeginUpdate()
            ' Apply a previously defined style to the table
            document.Tables(0).Style = tStyleMain
            document.EndUpdate()
        End Sub

        Private Sub deleteCellBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            table.BeginUpdate()
            ' Delete a cell
            table.Cell(1, 1).Delete()
            ' Delete a row
            table.Rows(0).Delete()
            table.EndUpdate()
        End Sub

        Private Sub deleteRowBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            table.BeginUpdate()
            ' Delete a row
            table.Rows(0).Delete()
            table.EndUpdate()
        End Sub

        Private Sub deleteColumnBtn_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            ' Call the declared method using ForEachRow method and the corresponding delegate
            table.ForEachRow(New TableRowProcessorDelegate(AddressOf DeleteCells))
        End Sub

        ' Declare a method that deletes the second cell in every table row
        Public Shared Sub DeleteCells(ByVal row As TableRow, ByVal i As Integer)
            row.Cells(1).Delete()
        End Sub

        Private Sub rotateButtonItem1_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            table.BeginUpdate()
            table(2, 0).TextDirection = TextDirection.Upward
            table.EndUpdate()
        End Sub
    End Class
End Namespace
