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
        const string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample Word document with headings and page breaks.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the first chapter content.");

        // Insert a page break to start a new chapter on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the second chapter content.");

        // -----------------------------------------------------------------
        // 2. Save the document as EPUB, splitting at heading paragraphs.
        // -----------------------------------------------------------------
        HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            Encoding = Encoding.UTF8,
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            ExportDocumentProperties = true
        };
        sourceDoc.Save(epubPath, epubSaveOptions);

        // Verify that the EPUB file was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("EPUB file was not created.");

        // -----------------------------------------------------------------
        // 3. Load the generated EPUB file.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // -----------------------------------------------------------------
        // 4. Convert the EPUB document to PDF.
        // -----------------------------------------------------------------
        epubDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");
    }
}
