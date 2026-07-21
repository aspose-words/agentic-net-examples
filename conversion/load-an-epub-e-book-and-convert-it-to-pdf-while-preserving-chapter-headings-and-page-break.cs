using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with headings and page breaks.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the introduction chapter. It contains some introductory text.");

        // Page break before next chapter.
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This chapter provides detailed information.");

        // Save the document as EPUB, splitting at heading paragraphs.
        string epubPath = "sample.epub";
        HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,
            Encoding = Encoding.UTF8,
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            ExportDocumentProperties = true
        };
        sourceDoc.Save(epubPath, epubSaveOptions);

        // Load the EPUB file.
        Document epubDoc = new Document(epubPath);

        // Convert the EPUB to PDF.
        string pdfPath = "output.pdf";
        epubDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF conversion failed; the output file was not created.");
    }
}
