using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with headings and page breaks.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Chapter 1 heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Chapter 1 content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 1.");

        // Insert a page break between chapters.
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2 heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Chapter 2 content.
        builder.Writeln("This is the content of chapter 2.");

        // Save the document as EPUB.
        const string epubPath = "input.epub";
        source.Save(epubPath, SaveFormat.Epub);

        // Load the EPUB file.
        Document epubDoc = new Document(epubPath);

        // Convert the EPUB to PDF.
        const string pdfPath = "output.pdf";
        epubDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF conversion failed; output file was not created.");
    }
}
