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

        // Insert a table with 3 rows and 3 columns.
        Table table = builder.StartTable();

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Apply the built‑in "Grid Table 5 Dark" style to the whole table.
        table.StyleIdentifier = StyleIdentifier.GridTable5Dark;

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
