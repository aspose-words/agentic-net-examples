using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // Build a table with 10 rows and 3 columns.
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Col {col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        Table table = builder.EndTable();

        // Apply alternating background colors to rows.
        for (int i = 0; i < table.Rows.Count; i++)
        {
            Color bgColor = (i % 2 == 0) ? Color.LightGray : Color.White;
            foreach (Cell cell in table.Rows[i].Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRows.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
