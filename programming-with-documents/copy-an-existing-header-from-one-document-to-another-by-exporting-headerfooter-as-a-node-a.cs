using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a source document with a header.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        srcBuilder.Writeln("Source Document Header");
        sourceDoc.Save("Source.docx");

        // Create a target document without a header (or with a default header).
        Document targetDoc = new Document();
        DocumentBuilder tgtBuilder = new DocumentBuilder(targetDoc);
        tgtBuilder.Writeln("Target document body.");
        targetDoc.Save("Target.docx");

        // Load the source document (already in memory) and get its primary header.
        HeaderFooter sourceHeader = sourceDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // Get the primary header of the target document.
        HeaderFooter targetHeader = targetDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // Clear any existing content in the target header.
        targetHeader.RemoveAllChildren();

        // Import each node from the source header into the target document and append it.
        foreach (Node node in sourceHeader)
        {
            Node importedNode = targetDoc.ImportNode(node, true, ImportFormatMode.KeepSourceFormatting);
            targetHeader.AppendChild(importedNode);
        }

        // Save the resulting document.
        string resultPath = "Result.docx";
        targetDoc.Save(resultPath);

        // Verify that the result file was created.
        if (File.Exists(resultPath))
        {
            Console.WriteLine($"Header copied successfully. Result saved to '{resultPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to create the result document.");
        }
    }
}
