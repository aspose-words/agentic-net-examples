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

        // Start a table with 3 columns.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.InsertCell();
        builder.Writeln("Cell 1,3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.InsertCell();
        builder.Writeln("Cell 2,3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Inform the user where the file was saved.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
