using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsNodeImportExample
{
    class Program
    {
        static void Main()
        {
            // Create temporary file paths.
            string tempDir = Path.GetTempPath();
            string sourcePath = Path.Combine(tempDir, "Source.docx");
            string destinationPath = Path.Combine(tempDir, "Destination.docx");
            string resultPath = Path.Combine(tempDir, "Result.docx");

            // -----------------------------------------------------------------
            // Build a source document containing a few tables.
            // -----------------------------------------------------------------
            Document srcDoc = new Document();
            Section srcSection = srcDoc.FirstSection ?? (Section)srcDoc.AppendChild(new Section(srcDoc));
            Body srcBody = srcSection.Body;

            // Table 0
            Table table0 = new Table(srcDoc);
            Row row0 = new Row(srcDoc);
            Cell cell0 = new Cell(srcDoc);
            Paragraph para0 = new Paragraph(srcDoc);
            para0.AppendChild(new Run(srcDoc, "Source Table 0, Cell 0"));
            cell0.AppendChild(para0);
            row0.AppendChild(cell0);
            table0.AppendChild(row0);
            srcBody.AppendChild(table0);

            // Some text between tables.
            Paragraph between = new Paragraph(srcDoc);
            between.AppendChild(new Run(srcDoc, "Between tables"));
            srcBody.AppendChild(between);

            // Table 1
            Table table1 = new Table(srcDoc);
            Row row1 = new Row(srcDoc);
            Cell cell1 = new Cell(srcDoc);
            Paragraph para1 = new Paragraph(srcDoc);
            para1.AppendChild(new Run(srcDoc, "Source Table 1, Cell 0"));
            cell1.AppendChild(para1);
            row1.AppendChild(cell1);
            table1.AppendChild(row1);
            srcBody.AppendChild(table1);

            // Save the source document.
            srcDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // Build a destination document (initially empty with a single paragraph).
            // -----------------------------------------------------------------
            Document dstDoc = new Document();
            Section dstSection = dstDoc.FirstSection ?? (Section)dstDoc.AppendChild(new Section(dstDoc));
            Body dstBody = dstSection.Body;
            Paragraph startPara = new Paragraph(dstDoc);
            startPara.AppendChild(new Run(dstDoc, "Destination document start"));
            dstBody.AppendChild(startPara);

            // Save the destination document.
            dstDoc.Save(destinationPath);

            // -----------------------------------------------------------------
            // Load the documents from the temporary files.
            // -----------------------------------------------------------------
            Document sourceDocument = new Document(sourcePath);
            Document destinationDocument = new Document(destinationPath);

            // Choose which tables to import (0‑based indices).
            int[] tablesToImport = { 0, 1 }; // Import both tables for demonstration.

            // Create a NodeImporter to handle style and list translation.
            NodeImporter importer = new NodeImporter(sourceDocument, destinationDocument, ImportFormatMode.KeepSourceFormatting);

            foreach (int tableIndex in tablesToImport)
            {
                // Retrieve the table from the source document.
                Table srcTable = (Table)sourceDocument.GetChild(NodeType.Table, tableIndex, true);
                if (srcTable == null)
                    continue; // Skip if the index is out of range.

                // Import the table node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(srcTable, true);

                // Append the imported table to the end of the destination document's body.
                destinationDocument.FirstSection.Body.AppendChild(importedNode);
            }

            // Save the modified destination document.
            destinationDocument.Save(resultPath);

            Console.WriteLine($"Result document saved to: {resultPath}");
        }
    }
}
