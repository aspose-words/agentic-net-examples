using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Build a 2x2 table where every cell's text is aligned to the bottom.
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Set vertical alignment for the current cell.
                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;

                // Write some sample text.
                builder.Write($"Row {row + 1}, Cell {col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellVerticalAlignmentBottom.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
