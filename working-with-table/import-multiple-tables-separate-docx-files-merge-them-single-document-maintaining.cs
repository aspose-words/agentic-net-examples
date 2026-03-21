using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeTablesExample
{
    static void Main()
    {
        // Create a temporary folder to hold the sample source documents.
        string tempFolder = Path.Combine(Path.GetTempPath(), "MergeTablesExample");
        Directory.CreateDirectory(tempFolder);

        // Generate three sample DOCX files, each containing a simple table.
        for (int i = 1; i <= 3; i++)
        {
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

            // Add a heading to identify the source document.
            srcBuilder.Writeln($"Source Document {i}");

            // Create a 2x2 table with sample data.
            Table table = srcBuilder.StartTable();
            srcBuilder.InsertCell();
            srcBuilder.Write($"Row 1, Cell 1 (Doc {i})");
            srcBuilder.InsertCell();
            srcBuilder.Write($"Row 1, Cell 2 (Doc {i})");
            srcBuilder.EndRow();

            srcBuilder.InsertCell();
            srcBuilder.Write($"Row 2, Cell 1 (Doc {i})");
            srcBuilder.InsertCell();
            srcBuilder.Write($"Row 2, Cell 2 (Doc {i})");
            srcBuilder.EndRow();

            srcBuilder.EndTable();

            // Save the source document.
            string srcPath = Path.Combine(tempFolder, $"Table{i}.docx");
            srcDoc.Save(srcPath);
        }

        // Create a blank destination document.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Optional title.
        builder.Writeln("Merged Tables:");
        // Keep a reference to the node after which new tables will be inserted.
        Node insertionPoint = builder.CurrentParagraph;

        // Paths to the source DOCX files that each contain a table.
        List<string> sourceFiles = new List<string>
        {
            Path.Combine(tempFolder, "Table1.docx"),
            Path.Combine(tempFolder, "Table2.docx"),
            Path.Combine(tempFolder, "Table3.docx")
        };

        foreach (string filePath in sourceFiles)
        {
            // Load the source document.
            Document srcDoc = new Document(filePath);

            // Iterate over every table in the source document.
            foreach (Table srcTable in srcDoc.GetChildNodes(NodeType.Table, true))
            {
                // Import the table into the destination document, preserving its original formatting.
                Node importedTable = dstDoc.ImportNode(srcTable, true, ImportFormatMode.KeepSourceFormatting);

                // Insert the imported table after the current insertion point.
                CompositeNode parent = (CompositeNode)insertionPoint.ParentNode;
                parent.InsertAfter(importedTable, insertionPoint);

                // Update the insertion point so the next table is placed after this one.
                insertionPoint = importedTable;

                // Insert a blank paragraph after each table for visual separation.
                Paragraph separator = new Paragraph(dstDoc);
                parent.InsertAfter(separator, insertionPoint);
                insertionPoint = separator;
            }
        }

        // Save the merged document in the temporary folder.
        string mergedPath = Path.Combine(tempFolder, "MergedTables.docx");
        dstDoc.Save(mergedPath);

        Console.WriteLine($"Merged document created at: {mergedPath}");
    }
}
