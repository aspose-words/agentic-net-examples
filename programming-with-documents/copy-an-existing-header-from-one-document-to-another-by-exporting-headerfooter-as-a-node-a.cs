using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string srcPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string dstPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");

        // -------------------------
        // Create the source document with a header.
        // -------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Move the builder into the primary header and add some text.
        srcBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        srcBuilder.Write("This is the source header.");

        // Return to the main body and add a paragraph.
        srcBuilder.MoveToSection(0);
        srcBuilder.Writeln("Body of the source document.");

        // Save the source document (required by the task to have an existing file).
        srcDoc.Save(srcPath);

        // -------------------------
        // Create the destination document.
        // -------------------------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Body of the destination document.");

        // -------------------------
        // Export the header from the source document as a node and import it into the destination.
        // -------------------------
        // Retrieve the primary header from the source.
        HeaderFooter srcHeader = srcDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // Import the header node into the destination document.
        // Use KeepSourceFormatting to preserve the original header appearance.
        Node importedHeaderNode = dstDoc.ImportNode(srcHeader, true, ImportFormatMode.KeepSourceFormatting);

        // Add the imported header to the destination's first section.
        dstDoc.FirstSection.HeadersFooters.Add((HeaderFooter)importedHeaderNode);

        // -------------------------
        // Save the destination document.
        // -------------------------
        dstDoc.Save(dstPath);
    }
}
