using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

namespace TableToImageExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and build a simple 2x2 table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Save the source document (optional, just to have a reference file).
            string sourcePath = "SourceTable.docx";
            doc.Save(sourcePath);

            // Retrieve the first table from the document.
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Create a new document that will contain only the extracted table.
            Document tableDoc = new Document();

            // Import the table node into the new document.
            // The NodeImporter constructor expects the source document, not the node itself.
            NodeImporter importer = new NodeImporter(doc, tableDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(table, true);
            tableDoc.FirstSection.Body.AppendChild(importedTable);

            // Render the document (which now contains only the table) to a PNG image.
            string imagePath = "TableImage.png";
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the first page where the table resides.
                PageSet = new PageSet(0)
            };
            tableDoc.Save(imagePath, options);

            // Verify that the image file was created.
            if (!File.Exists(imagePath))
                throw new InvalidOperationException($"Failed to create image file: {imagePath}");
        }
    }
}
