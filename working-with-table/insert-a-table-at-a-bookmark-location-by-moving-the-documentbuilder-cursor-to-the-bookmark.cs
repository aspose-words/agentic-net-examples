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

        // Insert some introductory text.
        builder.Writeln("Document with a bookmark where a table will be inserted.");

        // Insert a bookmark named "InsertHere".
        builder.StartBookmark("InsertHere");
        builder.Writeln("This paragraph is inside the bookmark.");
        builder.EndBookmark("InsertHere");

        // Move the builder to the bookmark location.
        builder.MoveToBookmark("InsertHere");

        // Insert a table at the bookmark location.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell A2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, you could load the document again to ensure it is readable.
        Document loadedDoc = new Document(outputPath);
        // No further actions; program ends here.
    }
}
