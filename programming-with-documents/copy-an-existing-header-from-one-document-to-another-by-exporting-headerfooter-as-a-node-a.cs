using System;
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
        // Save the source document (required for the example workflow).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Create a destination document (initially empty).
        Document destDoc = new Document();
        // Ensure the destination document has at least one section.
        // (A new Document already contains a single empty section.)

        // Export the header from the source document.
        HeaderFooter sourceHeader = sourceDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        // Import the header node into the destination document.
        HeaderFooter importedHeader = (HeaderFooter)destDoc.ImportNode(sourceHeader, true);
        // Add the imported header to the destination document's first section.
        destDoc.FirstSection.HeadersFooters.Add(importedHeader);

        // Save the destination document.
        string destPath = "Destination.docx";
        destDoc.Save(destPath);

        // Optional verification: reopen the destination document and read the header text.
        Document verifyDoc = new Document(destPath);
        HeaderFooter destHeader = verifyDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        string headerText = destHeader.GetText();
        // Output the header text to demonstrate successful copy (no user interaction required).
        Console.WriteLine("Copied Header Text: " + headerText.Trim());
    }
}
