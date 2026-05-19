using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a large table (e.g., 1000 rows × 5 columns).
        Table table = builder.StartTable();

        // Create header row.
        for (int col = 0; col < 5; col++)
        {
            builder.InsertCell();
            builder.Write($"Header {col + 1}");
        }
        builder.EndRow();

        // Populate the table with many rows.
        for (int row = 0; row < 1000; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Force a layout pass after all modifications.
        doc.UpdatePageLayout();

        // Save the document.
        string outputPath = "LargeTableOptimized.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
