using DevExpress.Office.Utils;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

namespace TablesSimpleExample
{
    public partial class Form1 : RibbonForm
    {
        Table table;
        Document document;
        public Form1()
        {
            InitializeComponent();
            document = richEditControl1.Document;
        }

        private void createTablebtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            // Create a new table and specify its layout type
            table = document.Tables.Create(document.Range.End, 2, 2);

            // Add new rows to the table
            table.Rows.InsertBefore(0);
            table.Rows.InsertAfter(0);

            // Add a new column to the table
            table.Rows[0].Cells.Append();
            table.Rows[0].Cells.InsertBefore(0);

            table.Rows[0].FirstCell.PreferredWidthType = WidthType.Auto;

            table.Rows[0].Cells[1].PreferredWidthType = WidthType.Fixed;
            table.Rows[0].Cells[1].PreferredWidth = Units.InchesToDocumentsF(0.8f);


            // Set the second column width and cell height
            table[0, 2].PreferredWidthType = WidthType.Fixed;
            table[0, 2].PreferredWidth = Units.InchesToDocumentsF(5f);
            table[0, 2].HeightType = HeightType.Exact;
            table[0, 2].Height = Units.InchesToDocumentsF(0.5f);

            //Set the third column width 
            table.Rows[0].LastCell.PreferredWidthType = WidthType.Fixed;
            table.Rows[0].LastCell.PreferredWidth = Units.InchesToDocumentsF(0.8f);
        }

        private void mergeBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            table.BeginUpdate();

            // Split cell into 8
            table[3, 2].Split(4, 2);

            // Merge cells
            table.MergeCells(table[4, 2], table[4, 3]);
            table.MergeCells(table[6, 2], table[6, 3]);
            table.MergeCells(table[2, 0], table[6, 0]);
            table.EndUpdate();
        }

        private void repeatRowsBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            table.BeginUpdate();

            // Repeat first three rows as header:
            table.Rows[0].RepeatAsHeaderRow = true;
            table.Rows[1].RepeatAsHeaderRow = true;
            table.Rows[2].RepeatAsHeaderRow = true;

            // Break last row across pages:
            table.LastRow.BreakAcrossPages = true;
            table.EndUpdate();
        }


        private void wrapTextBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            table.BeginUpdate();

            // Wrap text around the table
            table.TextWrappingType = TableTextWrappingType.Around;

            // Specify vertical alignment:
            table.RelativeVerticalPosition = TableRelativeVerticalPosition.Paragraph;
            table.VerticalAlignment = TableVerticalAlignment.None;
            table.OffsetYRelative = Units.InchesToDocumentsF(2f);

            // Specify horizontal alignment:
            table.RelativeHorizontalPosition = TableRelativeHorizontalPosition.Margin;
            table.HorizontalAlignment = TableHorizontalAlignment.Center;

            // Set distance between the text and the table:
            table.MarginBottom = Units.InchesToDocumentsF(0.3f);
            table.MarginLeft = Units.InchesToDocumentsF(0.3f);
            table.MarginTop = Units.InchesToDocumentsF(0.3f);
            table.MarginRight = Units.InchesToDocumentsF(0.3f);
            table.EndUpdate();
        }

        private void insertContentBtn_ItemClick(object sender, ItemClickEventArgs e)
        {

            // Insert header data
            document.InsertSingleLineText(table.Rows[0].Cells[2].Range.Start, "Active Customers");
            document.InsertSingleLineText(table[2, 1].Range.Start, "Photo");
            document.InsertSingleLineText(table[2, 0].Range.Start, "Customer №1");
            document.InsertSingleLineText(table[2, 2].Range.Start, "Customer Info");
            document.InsertSingleLineText(table[2, 3].Range.Start, "Rentals");

            // Insert the customer photo
            document.Images.Insert(table[3, 1].Range.Start, DocumentImageSource.FromFile("photo.png"));

            // Insert the customer info
            document.InsertText(table[3, 2].Range.Start, "Ryan Anita W");
            document.InsertText(table[3, 3].Range.Start, "Intermediate");
            document.InsertText(table[4, 2].Range.Start, "3/28/1984");
            document.InsertText(table[5, 2].Range.Start, "anita_ryan@dxvideorent.com");
            document.InsertText(table[5, 3].Range.Start, "(555)421-0059");
            document.InsertText(table[6, 2].Range.Start, "5119 Beryl Dr, San Antonio, TX 78212");

            document.InsertSingleLineText(table[3, 4].Range.Start, "18");
        }

        private void formatContentBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            // Apply formatting to the "Active Customers" cell
            CharacterProperties properties = document.BeginUpdateCharacters(table[0, 2].ContentRange);
            properties.FontName = "Segoe UI";
            properties.FontSize = 16;
            document.EndUpdateCharacters(properties);
            ParagraphProperties alignment = document.BeginUpdateParagraphs(table[0, 2].ContentRange);
            alignment.Alignment = ParagraphAlignment.Center;
            document.EndUpdateParagraphs(alignment);
            table[0, 1].VerticalAlignment = TableCellVerticalAlignment.Center;

            // Apply formatting to the header cells
            CharacterProperties headerRowProperties = document.BeginUpdateCharacters(table.Rows[2].Range);
            headerRowProperties.FontName = "Segoe UI";
            headerRowProperties.FontSize = 11;
            headerRowProperties.ForeColor = Color.FromArgb(212, 236, 183);
            document.EndUpdateCharacters(headerRowProperties);

            ParagraphProperties headerRowParagraphProperties = document.BeginUpdateParagraphs(table.Rows[2].Range);
            headerRowParagraphProperties.Alignment = ParagraphAlignment.Center;
            document.EndUpdateParagraphs(headerRowParagraphProperties);

            // Apply formatting to the customer info cells
            DocumentRange targetRange = document.CreateRange(table[3, 2].Range.Start, table[6, 3].Range.Start.ToInt() - table[3, 2].Range.Start.ToInt());
            CharacterProperties infoProperties = document.BeginUpdateCharacters(targetRange);
            infoProperties.FontSize = 8;
            infoProperties.FontName = "Segoe UI";
            infoProperties.ForeColor = Color.FromArgb(111, 116, 106);
            document.EndUpdateCharacters(infoProperties);

            // Format "Rentals" cells
            CharacterProperties rentalFormat = document.BeginUpdateCharacters(table[3, 4].Range);
            rentalFormat.FontSize = 28;
            rentalFormat.Bold = true;
            document.EndUpdateCharacters(rentalFormat);

            ParagraphProperties rentalAlignment = document.BeginUpdateParagraphs(table[3, 4].Range);
            rentalAlignment.Alignment = ParagraphAlignment.Center;
            document.EndUpdateParagraphs(rentalAlignment);
            table[3, 4].VerticalAlignment = TableCellVerticalAlignment.Center;
        }

        private void customizeBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            table.BeginUpdate();

            // Call the ChangeCellBorderColor method
            // for every cell in the first two rows
            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j < table.Rows[i].Cells.Count; j++)
                {
                    // Specify the border style and the background color for the header cells 
                    table[i, j].Borders.Bottom.LineStyle = TableBorderLineStyle.None;
                    table[i, j].Borders.Left.LineStyle = TableBorderLineStyle.None;
                    table[i, j].Borders.Right.LineStyle = TableBorderLineStyle.None;
                    table[i, j].Borders.Top.LineStyle = TableBorderLineStyle.None;
                    table[i, j].BackgroundColor = Color.Transparent;
                }
            }

            TableRow targetRow = table.Rows[2];
            targetRow.Cells[1].BackgroundColor = Color.FromArgb(99, 122, 110);
            targetRow.Cells[2].BackgroundColor = Color.FromArgb(99, 122, 110);
            targetRow.Cells[3].BackgroundColor = Color.FromArgb(99, 122, 110);
            table.EndUpdate();

        }

        private void tableStyleBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            document.BeginUpdate();
            // Create a new table style
            TableStyle tStyleMain = document.TableStyles.CreateNew();

            // Specify style options
            TableBorder insideHorizontalBorder = tStyleMain.TableBorders.InsideHorizontalBorder;
            insideHorizontalBorder.LineStyle = TableBorderLineStyle.Single;
            insideHorizontalBorder.LineColor = Color.White;

            TableBorder insideVerticalBorder = tStyleMain.TableBorders.InsideVerticalBorder;
            insideVerticalBorder.LineStyle = TableBorderLineStyle.Single;
            insideVerticalBorder.LineColor = Color.White;
            tStyleMain.CellBackgroundColor = Color.FromArgb(227, 238, 220);
            tStyleMain.Name = "MyTableStyle";

            // Add the style to the document collection
            document.TableStyles.Add(tStyleMain);

            // Create conditional styles (styles for specific table elements)         
            TableConditionalStyle myNewStyleForOddRows = tStyleMain.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddRowBanding);
            myNewStyleForOddRows.CellBackgroundColor = Color.FromArgb(196, 220, 182);

            TableConditionalStyle myNewStyleForBottomRightCell = tStyleMain.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.BottomRightCell);
            myNewStyleForBottomRightCell.CellBackgroundColor = Color.FromArgb(188, 214, 201);
            document.EndUpdate();

            document.BeginUpdate();

            // Apply a previously defined style to the table
            document.Tables[0].Style = tStyleMain;
            document.EndUpdate();

        }

        private void deleteCellBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            table.BeginUpdate();

            // Delete a cell
            table.Cell(1, 1).Delete();

            // Delete a row
            table.Rows[0].Delete();
            table.EndUpdate();
        }

        private void deleteRowBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            table.BeginUpdate();

            // Delete a row
            table.Rows[0].Delete();
            table.EndUpdate();

        }

        private void deleteColumnBtn_ItemClick(object sender, ItemClickEventArgs e)
        {
            // Call the declared method using ForEachRow method and the corresponding delegate
            table.ForEachRow(new TableRowProcessorDelegate(DeleteCells));
        }

        // Declare a method that deletes the second cell in every table row
        public static void DeleteCells(TableRow row, int i)
        {
            row.Cells[1].Delete();
        }

        private void rotateButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            table.BeginUpdate();
            table[2, 0].TextDirection = TextDirection.Upward;
            table.EndUpdate();

        }
    }
}




