using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare file paths
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string destinationPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");

        // -------------------------
        // Create source document with a header
        // -------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        // Move to the primary header of the first section and add text
        srcBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        srcBuilder.Writeln("This is the copied header");
        // Save the source document (optional, just for demonstration)
        sourceDoc.Save(sourcePath);

        // -------------------------
        // Create destination document (initially without a header)
        // -------------------------
        Document destDoc = new Document();
        // Ensure there is at least one section
        if (destDoc.Sections.Count == 0)
            destDoc.AppendChild(new Section(destDoc));

        // -------------------------
        // Export the header from source and import into destination
        // -------------------------
        // Get the primary header from the source document
        HeaderFooter sourceHeader = sourceDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        // Import the header node into the destination document
        Node importedHeader = destDoc.ImportNode(sourceHeader, true);
        // Add the imported header to the destination document's first section
        destDoc.FirstSection.HeadersFooters.Add(importedHeader as HeaderFooter);

        // -------------------------
        // Save the destination document
        // -------------------------
        destDoc.Save(destinationPath);
    }
}
