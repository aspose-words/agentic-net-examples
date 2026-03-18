using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeTablesExample
{
    static void Main()
    {
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
            "Table1.docx",
            "Table2.docx",
            "Table3.docx"
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

        // Save the merged document.
        dstDoc.Save("MergedTables.docx");
    }
}
