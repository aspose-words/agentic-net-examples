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

        // ---------- First table ----------
        Table firstTable = builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Write("First Table - Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("First Table - Row 1, Cell 2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("First Table - Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("First Table - Row 2, Cell 2");
        builder.EndRow();

        // Finish first table.
        builder.EndTable();

        // Insert an empty paragraph to separate the tables.
        builder.InsertParagraph();

        // ---------- Second table ----------
        Table secondTable = builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Write("Second Table - Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Second Table - Row 1, Cell 2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("Second Table - Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Second Table - Row 2, Cell 2");
        builder.EndRow();

        // Finish second table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "JoinedTables.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
