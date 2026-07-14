using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string destinationPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string mergedPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.docx");

        // -----------------------------------------------------------------
        // 1. Create a source document with two tables.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.Writeln("Source Document - Table 1");
        // First table
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("R1C1");
        srcBuilder.InsertCell();
        srcBuilder.Write("R1C2");
        srcBuilder.EndRow();
        srcBuilder.InsertCell();
        srcBuilder.Write("R2C1");
        srcBuilder.InsertCell();
        srcBuilder.Write("R2C2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        srcBuilder.Writeln(); // Add a paragraph between tables
        srcBuilder.Writeln("Source Document - Table 2");
        // Second table
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("A");
        srcBuilder.InsertCell();
        srcBuilder.Write("B");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Save the source document.
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a destination document with some initial content.
        // -----------------------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destinationDoc);
        dstBuilder.Writeln("Destination Document - Start");
        // Save the destination document (optional, just for completeness).
        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Import the first table from the source document into the destination.
        // -----------------------------------------------------------------
        // Locate the table to import (e.g., the first table).
        Table sourceTable = sourceDoc.FirstSection.Body.Tables[0];

        // Prepare a NodeImporter with the desired import format mode.
        NodeImporter importer = new NodeImporter(sourceDoc, destinationDoc, ImportFormatMode.KeepSourceFormatting);

        // Import the table node (deep copy, including its children).
        Node importedTableNode = importer.ImportNode(sourceTable, true);

        // Insert the imported table after the first paragraph of the destination.
        Paragraph insertionPoint = destinationDoc.FirstSection.Body.FirstParagraph;
        destinationDoc.FirstSection.Body.InsertAfter(importedTableNode, insertionPoint);

        // -----------------------------------------------------------------
        // 4. Save the merged document.
        // -----------------------------------------------------------------
        destinationDoc.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 5. Validation: ensure the merged file exists and contains the imported table.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document mergedDoc = new Document(mergedPath);
        int tableCount = mergedDoc.GetChildNodes(NodeType.Table, true).Count;

        // The merged document should contain at least one table (the imported one).
        if (tableCount < 1)
            throw new InvalidOperationException("Merged document does not contain any tables.");

        // Output a simple confirmation.
        Console.WriteLine($"Merged document created successfully with {tableCount} table(s).");
    }
}
