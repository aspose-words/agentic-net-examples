using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a source document with a primary header.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        srcBuilder.Write("This is the source header.");

        // Save the source document (optional, just to demonstrate persistence).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Create a destination document with some body content.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Body content of the destination document.");

        // Export the header from the source document.
        HeaderFooter sourceHeader = sourceDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // Import the header node into the destination document.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
        HeaderFooter importedHeader = (HeaderFooter)importer.ImportNode(sourceHeader, true);

        // Add the imported header to the destination document's first section.
        destDoc.FirstSection.HeadersFooters.Add(importedHeader);

        // Save the destination document which now contains the copied header.
        string destPath = "Destination.docx";
        destDoc.Save(destPath);
    }
}
