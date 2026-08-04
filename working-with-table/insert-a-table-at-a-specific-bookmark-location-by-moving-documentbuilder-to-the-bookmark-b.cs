using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = "TableAtBookmark.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content and a bookmark where the table will be inserted.
        builder.Writeln("Document start.");
        builder.StartBookmark("InsertTableHere");
        builder.Writeln("This paragraph is inside the bookmark.");
        builder.EndBookmark("InsertTableHere");
        builder.Writeln("Document end.");

        // Move the builder's cursor to the bookmark.
        bool moved = builder.MoveToBookmark("InsertTableHere");
        if (!moved)
            throw new InvalidOperationException("Bookmark 'InsertTableHere' not found.");

        // Build a 2x2 table at the bookmark location.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
