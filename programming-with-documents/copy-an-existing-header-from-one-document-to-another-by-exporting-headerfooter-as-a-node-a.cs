using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string srcPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string dstPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");

        // -------------------------------------------------
        // Create the source document with a primary header.
        // -------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        srcBuilder.Write("Source Header");
        srcDoc.Save(srcPath);

        // -------------------------------------------------
        // Create the destination document with some body text.
        // -------------------------------------------------
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Destination body text.");
        dstDoc.Save(dstPath);

        // -------------------------------------------------
        // Export the header from the source document.
        // -------------------------------------------------
        HeaderFooter srcHeader = srcDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // -------------------------------------------------
        // Import the header into the destination document.
        // -------------------------------------------------
        // Get (or create) the destination header of the same type.
        HeaderFooter dstHeader = dstDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (dstHeader == null)
        {
            dstHeader = new HeaderFooter(dstDoc, HeaderFooterType.HeaderPrimary);
            dstDoc.FirstSection.HeadersFooters.Add(dstHeader);
        }

        // Use NodeImporter to import each child node of the source header.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Node child in srcHeader)
        {
            // Import the child node (paragraphs, tables, etc.) into the destination document.
            Node importedChild = importer.ImportNode(child, true);
            dstHeader.AppendChild(importedChild);
        }

        // -------------------------------------------------
        // Save the resulting document.
        // -------------------------------------------------
        dstDoc.Save(resultPath);
    }
}
