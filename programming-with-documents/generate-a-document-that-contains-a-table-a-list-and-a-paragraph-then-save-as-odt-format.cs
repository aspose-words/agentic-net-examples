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

        // Insert a paragraph.
        builder.Writeln("This is a sample paragraph added to the document.");

        // Insert a bulleted list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First list item");
        builder.Writeln("Second list item");
        builder.Writeln("Third list item");
        builder.ListFormat.RemoveNumbers();

        // Insert a table with 3 rows and 3 columns.
        Table table = builder.StartTable();

        // First row (header)
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Header {i + 1}");
        }
        builder.EndRow();

        // Data rows
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {row + 1}, Col {col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Define output file path.
        string outputPath = "output.odt";

        // Save the document in ODT format.
        doc.Save(outputPath, SaveFormat.Odt);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Failed to create the ODT file.");
        }
    }
}
