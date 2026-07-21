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
        builder.StartTable();

        // First row, two cells.
        builder.InsertCell();
        builder.Write("First table, Cell 1");
        builder.InsertCell();
        builder.Write("First table, Cell 2");
        builder.EndRow();

        // Second row, two cells.
        builder.InsertCell();
        builder.Write("First table, Cell 3");
        builder.InsertCell();
        builder.Write("First table, Cell 4");
        builder.EndTable();

        // Insert an empty paragraph to separate the tables.
        // This prevents Word from automatically merging the two tables.
        builder.Writeln();

        // ---------- Second table ----------
        builder.StartTable();

        // First row, two cells.
        builder.InsertCell();
        builder.Write("Second table, Cell 1");
        builder.InsertCell();
        builder.Write("Second table, Cell 2");
        builder.EndRow();

        // Second row, two cells.
        builder.InsertCell();
        builder.Write("Second table, Cell 3");
        builder.InsertCell();
        builder.Write("Second table, Cell 4");
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablesWithSeparator.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
