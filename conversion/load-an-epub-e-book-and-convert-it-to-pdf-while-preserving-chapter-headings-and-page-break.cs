using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with chapter headings and page breaks.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the first chapter.");
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the second chapter.");
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the third chapter.");

        // Save the document as EPUB.
        string epubPath = "sample.epub";
        sampleDoc.Save(epubPath, SaveFormat.Epub);

        // Load the EPUB file.
        Document epubDoc = new Document(epubPath);

        // Convert the EPUB to PDF.
        string pdfPath = "output.pdf";
        epubDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional: clean up intermediate EPUB file.
        if (File.Exists(epubPath))
        {
            File.Delete(epubPath);
        }
    }
}
