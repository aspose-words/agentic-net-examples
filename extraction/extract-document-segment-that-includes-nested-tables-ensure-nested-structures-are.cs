using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace AsposeWordsNestedTableExtraction
{
    class Program
    {
        static void Main()
        {
            // Load the source document that contains nested tables.
            Document srcDoc = new Document("SourceWithNestedTables.docx");

            // Find the first table that contains nested tables.
            // This example simply takes the first table in the document.
            // In a real scenario you could locate a specific table by bookmark, index, etc.
            Table outerTable = (Table)srcDoc.GetChild(NodeType.Table, 0, true);
            if (outerTable == null)
                throw new InvalidOperationException("No table found in the source document.");

            // Create a new empty document that will hold the extracted segment.
            Document dstDoc = new Document();

            // Use NodeImporter to import the outer table (with its nested tables) into the new document.
            // KeepSourceFormatting preserves the original formatting of the imported nodes.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(outerTable, true);

            // Insert the imported table into the destination document using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.InsertNode(importedTable);

            // Save the new document. The nested table structure is retained.
            dstDoc.Save("ExtractedSegment.docx");
        }
    }
}
