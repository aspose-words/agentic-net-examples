using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a source document with a table.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source Document Intro");
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Extract the first table from the source document.
        Table extractedTable = sourceDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (extractedTable == null)
            throw new InvalidOperationException("No table found in the source document.");

        // Create a destination document with a custom bookmark.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Document Start");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Placeholder paragraph for insertion");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination Document End");

        // Move to the bookmark.
        destBuilder.MoveToBookmark("InsertHere");

        // Import the table from the source document into the destination document.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(extractedTable, true);

        // Insert the imported table after the paragraph that contains the bookmark.
        Paragraph placeholderParagraph = destBuilder.CurrentParagraph;
        if (placeholderParagraph == null)
            throw new InvalidOperationException("Failed to locate the placeholder paragraph.");

        CompositeNode parent = placeholderParagraph.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Placeholder paragraph does not have a valid parent.");

        parent.InsertAfter(importedTable, placeholderParagraph);
        // Optionally remove the placeholder paragraph.
        placeholderParagraph.Remove();

        // Save the destination document.
        const string destPath = "destination.docx";
        destDoc.Save(destPath);

        // Write a simple JSON report about the extraction.
        var report = new
        {
            ExtractedNodeType = "Table",
            ExtractedNodeCount = 1,
            SourceDocument = sourcePath,
            DestinationDocument = destPath
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        const string reportPath = "extraction_report.json";
        File.WriteAllText(reportPath, jsonReport);

        // Validate that the output files were created.
        if (!File.Exists(destPath))
            throw new InvalidOperationException("Destination document was not created.");
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Extraction report was not created.");
    }
}
