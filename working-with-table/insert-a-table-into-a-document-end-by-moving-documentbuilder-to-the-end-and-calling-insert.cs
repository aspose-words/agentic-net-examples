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

        // Move the builder cursor to the end of the document.
        builder.MoveToDocumentEnd();

        // Start a table at the current position.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // ---- Second row ----
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // Finish the table. This moves the cursor just after the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "OutputTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Indicate successful completion (optional).
        Console.WriteLine("Document saved to: " + Path.GetFullPath(outputPath));
    }
}
