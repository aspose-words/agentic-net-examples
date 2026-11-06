using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table.
        Table table = builder.StartTable();

        // Add header row.
        for (int col = 1; col <= 3; col++)
        {
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Writeln($"Header {col}");
        }
        builder.EndRow();

        // Add data rows.
        for (int row = 1; row <= 5; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {row}, Col {col}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Apply alternating row shading.
        // Row 0 is the header; start shading from the first data row (index 1).
        for (int i = 1; i < table.Rows.Count; i++)
        {
            Row currentRow = table.Rows[i];
            Color shade = (i % 2 == 0) ? Color.LightGray : Color.White;

            foreach (Cell cell in currentRow.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = shade;
            }
        }

        // Save the document.
        string outputPath = "TableWithAlternatingShading.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
