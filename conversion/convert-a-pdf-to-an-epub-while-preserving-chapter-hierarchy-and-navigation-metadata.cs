using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF and final EPUB files.
        const string pdfPath = "sample.pdf";
        const string epubPath = "output.epub";

        // -------------------------------------------------
        // Step 1: Create a sample document with chapter headings.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Save the document as PDF – this will be the input for conversion.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // -------------------------------------------------
        // Step 2: Load the PDF file.
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -------------------------------------------------
        // Step 3: Convert the loaded PDF to EPUB, preserving chapter hierarchy.
        // -------------------------------------------------
        pdfDoc.Save(epubPath, SaveFormat.Epub);

        // Verify that the EPUB was created and is not empty.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("The EPUB file was not created.");

        FileInfo epubInfo = new FileInfo(epubPath);
        if (epubInfo.Length == 0)
            throw new InvalidOperationException("The EPUB file is empty.");
    }
}
