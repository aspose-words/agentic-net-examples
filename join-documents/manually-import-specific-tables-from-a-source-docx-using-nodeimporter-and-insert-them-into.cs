using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Create a source document that contains two tables.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First table.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 1 - Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 1 - Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Add a paragraph between tables.
        srcBuilder.Writeln();

        // Second table.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 2 - Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 2 - Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a destination document with some initial text.
        // -----------------------------------------------------------------
        string destinationPath = Path.Combine(artifactsDir, "Destination.docx");
        Document destinationDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destinationDoc);
        dstBuilder.Writeln("Destination document start.");
        destinationDoc.Save(destinationPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents from disk.
        // -----------------------------------------------------------------
        Document src = new Document(sourcePath);
        Document dst = new Document(destinationPath);

        // -----------------------------------------------------------------
        // Choose the specific table(s) to import.
        // For this example we import the first table only.
        // -----------------------------------------------------------------
        Table? tableToImport = src.FirstSection?.Body?.Tables?.Count > 0 ? src.FirstSection.Body.Tables[0] : null;
        if (tableToImport == null)
            throw new InvalidOperationException("Source document does not contain the expected table.");

        // -----------------------------------------------------------------
        // Import the table node using NodeImporter.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(src, dst, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(tableToImport, true);

        // -----------------------------------------------------------------
        // Insert the imported table into the destination document after the first paragraph.
        // -----------------------------------------------------------------
        Paragraph insertionParagraph = dst.FirstSection.Body.FirstParagraph;
        insertionParagraph.ParentNode.InsertAfter(importedTable, insertionParagraph);

        // -----------------------------------------------------------------
        // Save the merged document.
        // -----------------------------------------------------------------
        string mergedPath = Path.Combine(artifactsDir, "Merged.docx");
        dst.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the merged file exists and contains at least one table.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not created.", mergedPath);

        Document mergedDoc = new Document(mergedPath);
        int tableCount = mergedDoc.GetChildNodes(NodeType.Table, true).Count;
        if (tableCount == 0)
            throw new InvalidOperationException("Merged document does not contain any tables.");

        // Output a simple confirmation to the console.
        Console.WriteLine($"Merged document created successfully at: {mergedPath}");
        Console.WriteLine($"Number of tables in merged document: {tableCount}");
    }
}
