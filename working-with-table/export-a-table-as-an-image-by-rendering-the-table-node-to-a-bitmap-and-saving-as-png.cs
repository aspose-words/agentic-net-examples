using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;   // Needed for the Table class

namespace TableToImageExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a source document and build a simple 2x2 table.
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

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

            // Create a destination document that will contain only the table.
            Document destDoc = new Document();

            // Import the table node from the source document into the destination document.
            NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(table, true);
            destDoc.FirstSection.Body.AppendChild(importedTable);

            // Save the destination document as a PNG image (rendering only the first page which holds the table).
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(0) // zero‑based page index
            };
            destDoc.Save("TableImage.png", options);
        }
    }
}
