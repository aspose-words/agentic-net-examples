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

        // ---------- First Table ----------
        builder.StartTable();
        builder.InsertCell();
        builder.Write("First table, Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Insert an empty paragraph to separate the tables.
        builder.InsertParagraph();

        // ---------- Second Table ----------
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Second table, Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "JoinedTables.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
