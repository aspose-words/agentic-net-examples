using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string epubPath = "sample.epub";
        const string pdfPath = "sample.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample document with headings and page breaks.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the first chapter content.");
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the second chapter content.");
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3: Conclusion");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the final chapter content.");

        // -----------------------------------------------------------------
        // Step 2: Save the document as EPUB, splitting at heading paragraphs.
        // -----------------------------------------------------------------
        HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,
            Encoding = Encoding.UTF8,
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            ExportDocumentProperties = true
        };
        sourceDoc.Save(epubPath, epubSaveOptions);

        // Verify that the EPUB file was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("EPUB file was not created.");

        // -----------------------------------------------------------------
        // Step 3: Load the EPUB and convert it to PDF.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);
        epubDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Optional: Inform that conversion succeeded.
        Console.WriteLine("EPUB successfully converted to PDF.");
    }
}
