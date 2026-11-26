using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure cell formatting: borders and uniform padding (5 points on each side).
        builder.CellFormat.Borders.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Color = Color.Black;
        builder.CellFormat.LeftPadding = 5;
        builder.CellFormat.RightPadding = 5;
        builder.CellFormat.TopPadding = 5;
        builder.CellFormat.BottomPadding = 5;

        // Build a complex outer table (3 rows x 3 columns).
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.Writeln($"R{row + 1}C{col + 1}");

                // End the row after the last cell in the current row.
                if (col == 2)
                    builder.EndRow();
            }
        }
        builder.EndTable();

        // Retrieve the outer table.
        Table outerTable = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Insert a nested table into the first cell of the outer table.
        Cell nestedCell = outerTable.Rows[0].Cells[0];
        builder.MoveTo(nestedCell.FirstParagraph);

        // Build nested table (2 rows x 2 columns).
        builder.StartTable();
        for (int nRow = 0; nRow < 2; nRow++)
        {
            for (int nCol = 0; nCol < 2; nCol++)
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.Writeln($"N{nRow + 1}C{nCol + 1}");

                // End the row after the last cell in the current nested row.
                if (nCol == 1)
                    builder.EndRow();
            }
        }
        builder.EndTable();

        // Save the sample DOCX.
        string docxPath = "sample.docx";
        doc.Save(docxPath);

        // Load the saved document (simulating a real conversion scenario).
        Document loadedDoc = new Document(docxPath);

        // Save as HTML, preserving borders and cell padding.
        string htmlPath = "output.html";
        loadedDoc.Save(htmlPath, SaveFormat.Html);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML conversion failed: output file not found.");
    }
}
