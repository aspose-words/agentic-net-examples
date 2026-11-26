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
            throw new Exception("Table was not created.");

        // Create a new empty document that will contain only the imported table.
        Document tableDoc = new Document();

        // Import the table node into the new document (required for cross‑document cloning).
        NodeImporter importer = new NodeImporter(doc, tableDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(table, true);

        // Append the imported table to the body of the new document.
        tableDoc.FirstSection.Body.AppendChild(importedTable);

        // Render the table document to a PNG image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Save only the first page (the table occupies a single page).
            PageSet = new PageSet(0)
        };

        string imagePath = "TableImage.png";
        tableDoc.Save(imagePath, options);

        // Verify that the image file was created.
        if (!File.Exists(imagePath))
            throw new Exception("Failed to save the table image.");
    }
}
