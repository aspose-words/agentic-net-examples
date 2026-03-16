using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsNodeImportExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source and destination documents.
            string sourcePath = @"C:\Docs\Source.docx";
            string destinationPath = @"C:\Docs\Destination.docx";
            string resultPath = @"C:\Docs\Result.docx";

            // Load the source and destination documents using the Document(string) constructor.
            Document srcDoc = new Document(sourcePath);
            Document dstDoc = new Document(destinationPath);

            // Choose which tables to import.
            // Example: import the first and third tables (0‑based index).
            // Adjust the indices or selection logic as needed.
            int[] tablesToImport = { 0, 2 };

            // Create a NodeImporter that will handle style and list translation.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            // Iterate over the selected table indices.
            foreach (int tableIndex in tablesToImport)
            {
                // Retrieve the table from the source document.
                Table srcTable = (Table)srcDoc.GetChild(NodeType.Table, tableIndex, true);
                if (srcTable == null)
                    continue; // Skip if the index is out of range.

                // Import the table node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(srcTable, true);

                // Append the imported table to the end of the destination document's body.
                // You can change the insertion point as required (e.g., InsertAfter a specific node).
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            // Save the modified destination document using the Save(string) method.
            dstDoc.Save(resultPath);
        }
    }
}
