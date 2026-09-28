using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with 3 rows and 3 columns.
        Table table = builder.StartTable();

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                // Insert cell text.
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            // End the current row.
            builder.EndRow();
        }

        // End the table construction.
        builder.EndTable();

        // Apply the built‑in "Grid Table 5 Dark" style to the entire table.
        table.StyleIdentifier = StyleIdentifier.GridTable5Dark;

        // Define output path.
        string outputPath = "Output.docx";

        // Save the document.
        doc.Save(outputPath);

        // Optional: verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
