using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class ExportTableAsImage
{
    public static void Main()
    {
        // Create a new blank document and build a simple 2x2 table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();
        builder.EndTable();

        // Get the first table from the source document.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Create a new document that will contain only the imported table.
        Document tableDoc = new Document();

        // Import the table node into the new document using NodeImporter.
        NodeImporter importer = new NodeImporter(firstTable.Document, tableDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(firstTable, true);
        tableDoc.FirstSection.Body.AppendChild(importedTable);

        // Render the document (which now contains only the table) to a PNG image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page (the table occupies a single page).
            PageSet = new PageSet(0)
        };

        string outputPath = "TableImage.png";
        tableDoc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the PNG image.");
    }
}
