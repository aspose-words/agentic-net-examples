using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string destinationPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");

        // ---------- Create source document with a header ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // Move cursor to the primary header and write header text.
        srcBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        srcBuilder.Write("This is the source header.");

        // Add some body content to make the document non‑empty.
        srcBuilder.MoveToDocumentEnd();
        srcBuilder.Writeln("Body of the source document.");

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // ---------- Create destination document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Add body content to the destination document.
        destBuilder.Writeln("Body of the destination document.");

        // ---------- Export the header from source and import into destination ----------
        // Retrieve the primary header from the source document.
        HeaderFooter sourceHeader = sourceDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // Import the header node into the destination document.
        HeaderFooter importedHeader = (HeaderFooter)destDoc.ImportNode(sourceHeader, true);

        // Add the imported header to the destination document's first section.
        destDoc.FirstSection.HeadersFooters.Add(importedHeader);

        // Save the destination document with the copied header.
        destDoc.Save(destinationPath);
    }
}
