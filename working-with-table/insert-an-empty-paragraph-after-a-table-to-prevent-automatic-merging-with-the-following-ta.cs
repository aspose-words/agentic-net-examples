using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the first table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1, Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Insert an empty paragraph to separate the tables.
        // This prevents Aspose.Words from merging the two tables automatically.
        builder.Writeln();

        // Build the second table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2, Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
