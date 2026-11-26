using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with headings and page breaks.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Chapter 1 heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Chapter 1 content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 1.");

        // Page break between chapters.
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2 heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Chapter 2 content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 2.");

        // Save the document as EPUB (input for conversion).
        const string epubPath = "sample.epub";
        sampleDoc.Save(epubPath, SaveFormat.Epub);

        // Load the EPUB file.
        Document epubDoc = new Document(epubPath);

        // Convert EPUB to PDF.
        const string pdfPath = "output.pdf";
        epubDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional: clean up sample files (comment out if you want to keep them).
        // File.Delete(epubPath);
        // File.Delete(pdfPath);
    }
}
