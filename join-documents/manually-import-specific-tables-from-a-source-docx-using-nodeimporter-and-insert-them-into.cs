using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        string sourcePath = "Source.docx";
        string destinationPath = "Destination.docx";
        string mergedPath = "Merged.docx";

        // ---------------------------------------------------------------
        // Create a source document containing two tables with distinct text.
        // ---------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First table.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 1 - Cell A1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 1 - Cell B1");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Add a paragraph between tables to separate them.
        srcBuilder.Writeln();

        // Second table.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 2 - Cell A1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Table 2 - Cell B1");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Save the source document (required by the rules).
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // Create a destination document with a placeholder paragraph.
        // ---------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destDoc);
        dstBuilder.Writeln("Destination document start.");
        // Save the destination document for completeness.
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // Import specific tables from the source document into the destination.
        // ---------------------------------------------------------------
        // Retrieve all tables from the source document.
        NodeCollection sourceTables = sourceDoc.GetChildNodes(NodeType.Table, true);

        // Choose which tables to import (e.g., both tables in this example).
        int[] tablesToImport = { 0, 1 };

        // Locate the insertion point – the first paragraph in the destination.
        Paragraph insertionParagraph = destDoc.FirstSection.Body.FirstParagraph;

        // Prepare the NodeImporter with the desired formatting mode.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert each selected table after the insertion paragraph.
        Node currentInsertionNode = insertionParagraph;
        foreach (int tableIndex in tablesToImport)
        {
            if (tableIndex < 0 || tableIndex >= sourceTables.Count)
                continue; // Safety check.

            Table srcTable = (Table)sourceTables[tableIndex];
            // Import the table node (deep copy) into the destination document.
            Node importedTable = importer.ImportNode(srcTable, true);
            // Insert the imported table after the current insertion node.
            CompositeNode parent = currentInsertionNode.ParentNode;
            parent.InsertAfter(importedTable, currentInsertionNode);
            // Update the insertion point so the next table is placed after this one.
            currentInsertionNode = importedTable;
        }

        // Save the merged document.
        destDoc.Save(mergedPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // Validation: ensure the merged file exists and contains expected text.
        // ---------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException($"Merged file was not created at '{mergedPath}'.");

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("Table 1 - Cell A1") || !mergedText.Contains("Table 2 - Cell A1"))
            throw new InvalidOperationException("Merged document does not contain the expected table contents.");

        // Successful execution – no console output required.
    }
}
