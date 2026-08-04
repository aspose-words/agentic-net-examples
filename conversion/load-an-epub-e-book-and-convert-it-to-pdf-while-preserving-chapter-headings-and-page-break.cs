using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string epubPath = "sample.epub";
        const string pdfPath = "sample.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample Word document with headings and page breaks.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1 (Heading 1)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.InsertBreak(BreakType.PageBreak);

        // Section 1.1 (Heading 2)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Content of section 1.1.");

        // Section 1.2 (Heading 2)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2");
        builder.Writeln("Content of section 1.2.");
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2 (Heading 1)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.InsertBreak(BreakType.PageBreak);

        // Section 2.1 (Heading 2)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");
        builder.Writeln("Content of section 2.1.");

        // -----------------------------------------------------------------
        // Step 2: Save the document as an EPUB file.
        // Use HtmlSaveOptions to configure EPUB output.
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
            throw new InvalidOperationException($"EPUB file '{epubPath}' was not created.");

        // -----------------------------------------------------------------
        // Step 3: Load the EPUB file.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(epubPath);

        // -----------------------------------------------------------------
        // Step 4: Convert the loaded EPUB to PDF.
        // -----------------------------------------------------------------
        epubDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"PDF file '{pdfPath}' was not created.");

        // The example finishes without waiting for user input.
    }
}
