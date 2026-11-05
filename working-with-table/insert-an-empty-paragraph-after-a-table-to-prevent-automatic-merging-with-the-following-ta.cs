using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the first table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1, Cell 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Table 1, Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Insert an empty paragraph to separate the tables.
        builder.Writeln();

        // Build the second table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2, Cell 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Table 2, Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Indicate successful creation.
        Console.WriteLine("Document created: " + Path.GetFullPath(outputPath));
    }
}
