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

        // Start building a table.
        Table table = builder.StartTable();

        // First row with two cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row with two cells.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a custom left indent (margin) to the table.
        table.LeftIndent = 30; // points

        // Table.RightIndent does not exist in this version of Aspose.Words.
        // Use DistanceRight to achieve a similar right‑margin effect.
        table.DistanceRight = 30; // points

        // Save the document to the local file system.
        const string outputPath = "TableMargin.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!System.IO.File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
