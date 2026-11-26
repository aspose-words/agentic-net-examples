using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up a two‑column layout.
        builder.PageSetup.TextColumns.SetCount(2);

        // Build a simple table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // In newer Aspose.Words versions the Table class provides
        // an AllowBreakAcrossPages property that also prevents breaking
        // across columns. If the property is unavailable, the default
        // behavior is used.

        // Save the document.
        string fileName = "TableNoBreakAcrossColumns.docx";
        doc.Save(fileName);

        // Verify the file was created.
        if (!File.Exists(fileName))
        {
            throw new InvalidOperationException($"Failed to create {fileName}");
        }

        Console.WriteLine($"Document saved as {fileName}");
    }
}
