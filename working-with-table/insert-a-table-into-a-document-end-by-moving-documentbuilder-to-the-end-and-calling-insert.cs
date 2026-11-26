using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to the end of the document.
        builder.MoveToDocumentEnd();

        // Build a table with 2 rows and 3 columns.
        builder.StartTable();

        // First row cells.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Writeln($"R1C{col + 1}");
        }
        builder.EndRow();

        // Second row cells.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Writeln($"R2C{col + 1}");
        }
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "TableAtEnd.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine("Document created successfully.");
    }
}
