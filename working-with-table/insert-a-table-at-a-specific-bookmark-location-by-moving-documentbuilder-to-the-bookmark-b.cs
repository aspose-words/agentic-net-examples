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

        // Add some introductory text.
        builder.Writeln("Document before the bookmark.");

        // Insert an empty bookmark where the table will be placed.
        builder.StartBookmark("TableLocation");
        builder.EndBookmark("TableLocation");

        // Add more text after the bookmark to demonstrate positioning.
        builder.Writeln("Document after the bookmark.");

        // Move the builder's cursor to the bookmark.
        bool moved = builder.MoveToBookmark("TableLocation");
        if (!moved)
            throw new InvalidOperationException("Bookmark not found.");

        // Build a 2x2 table at the bookmark location.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Define output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableAtBookmark.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
