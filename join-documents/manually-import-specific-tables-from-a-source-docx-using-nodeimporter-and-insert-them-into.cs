using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for sample files
        string sourcePath = "Source.docx";
        string destinationPath = "Destination.docx";
        string mergedPath = "Merged.docx";

        // -----------------------------------------------------------------
        // Create source document with multiple tables
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.Writeln("Source Document - Introduction");

        // Table 1
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 1 - Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 1 - Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        srcBuilder.Writeln(); // spacer

        // Table 2
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 2 - Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 2 - Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        srcBuilder.Writeln(); // spacer

        // Table 3
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 3 - Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 3 - Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create destination document with some initial content
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Destination Document - Header");
        destBuilder.Writeln("This is the original content of the destination document.");
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load both documents (simulating real scenario)
        // -----------------------------------------------------------------
        Document src = new Document(sourcePath);
        Document dst = new Document(destinationPath);

        // Prepare a NodeImporter to import nodes from source to destination
        NodeImporter importer = new NodeImporter(src, dst, ImportFormatMode.KeepSourceFormatting);

        // Retrieve all tables from the source document
        NodeCollection sourceTables = src.GetChildNodes(NodeType.Table, true);

        // Define which tables to import (e.g., first two tables)
        int tablesToImport = Math.Min(2, sourceTables.Count);

        // Insert imported tables after the first paragraph in the destination document
        Node insertionPoint = dst.FirstSection.Body.FirstParagraph;

        for (int i = 0; i < tablesToImport; i++)
        {
            Table srcTable = (Table)sourceTables[i];
            // Import the table node into the destination document
            Node importedNode = importer.ImportNode(srcTable, true);
            // Insert the imported table after the current insertion point
            insertionPoint = dst.FirstSection.Body.InsertAfter(importedNode, insertionPoint);
        }

        // Save the merged document
        dst.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
        {
            throw new InvalidOperationException($"Merged document was not created at path: {mergedPath}");
        }

        Document mergedDoc = new Document(mergedPath);
        int mergedTableCount = mergedDoc.GetChildNodes(NodeType.Table, true).Count;
        int expectedTableCount = tablesToImport + dst.GetChildNodes(NodeType.Table, true).Count - tablesToImport; // original tables in destination (0) + imported tables

        if (mergedTableCount != expectedTableCount)
        {
            throw new InvalidOperationException($"Merged document contains {mergedTableCount} tables, but expected {expectedTableCount}.");
        }

        // Optional cleanup (commented out)
        // File.Delete(sourcePath);
        // File.Delete(destinationPath);
        // File.Delete(mergedPath);
    }
}
