using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with chapter headings.
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

        // Save the document as PDF (input for conversion).
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF file.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to EPUB while preserving headings as navigation.
        string epubPath = "output.epub";
        pdfDoc.Save(epubPath, SaveFormat.Epub);

        // Verify that the EPUB file was created.
        if (!File.Exists(epubPath))
        {
            throw new InvalidOperationException("EPUB file was not created.");
        }
    }
}
