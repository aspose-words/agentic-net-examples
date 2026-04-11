using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary EPUB and the resulting PDF.
        string epubPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.epub");
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.pdf");

        // -------------------------------------------------
        // 1. Create a sample document containing headings and page breaks.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First chapter (heading level 1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of the first chapter.");

        // Page break to start a new chapter on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Second chapter (heading level 1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of the second chapter.");

        // Save the document as an EPUB file – this will be the input for conversion.
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Verify that the EPUB file was created successfully.
        if (!File.Exists(epubPath) || new FileInfo(epubPath).Length == 0)
            throw new InvalidOperationException("Failed to create the EPUB file.");

        // -------------------------------------------------
        // 2. Load the EPUB e‑book.
        // -------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // -------------------------------------------------
        // 3. Convert the EPUB to PDF while preserving headings in the PDF outline.
        // -------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Include headings up to level 3 in the PDF outline (table of contents).
            OutlineOptions = { HeadingsOutlineLevels = 3 }
        };

        epubDoc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created successfully.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF conversion failed.");

        // Indicate successful completion.
        Console.WriteLine("EPUB successfully converted to PDF.");
    }
}
