using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
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

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // In the current Aspose.Words version the AllowOverlap property is read‑only
        // and overlapping is enabled by default, so no explicit assignment is required.
        // If a future version provides a setter, the following line could be used:
        // table.AllowOverlap = true;

        // Save the document.
        string outputPath = "TableAllowOverlap.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Output a simple confirmation.
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
