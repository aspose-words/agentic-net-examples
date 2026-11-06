using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and build a simple 2‑cell table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new Exception("No table found in the document.");

        // Clone the table into a separate document.
        Document tableDoc = new Document();               // New empty document.
        Section section = tableDoc.Sections[0];           // Existing section to host the table.

        // Import the table into the new document before appending.
        NodeImporter importer = new NodeImporter(table.Document, tableDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(table, true);
        section.Body.AppendChild(importedTable);

        // Configure image save options to produce a PNG of the first page (where the table resides).
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0)   // Render only the first page.
        };

        string outputPath = "TableImage.png";
        tableDoc.Save(outputPath, options);

        // Verify that the image file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the PNG image of the table.");

        Console.WriteLine($"Table exported successfully as image: {Path.GetFullPath(outputPath)}");
    }
}
