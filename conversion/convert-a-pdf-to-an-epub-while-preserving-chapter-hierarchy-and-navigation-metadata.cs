using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "sample.pdf";
        const string epubPath = "output.epub";

        // -----------------------------------------------------------------
        // Step 1: Create a sample Word document with heading styles.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1 (Heading 1)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Section 1.1 (Heading 2)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Content of section 1.1.");

        // Subsection 1.1.1 (Heading 3)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1");
        builder.Writeln("Content of subsection 1.1.1.");

        // Chapter 2 (Heading 1)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.Writeln("Content of chapter 2.");

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to EPUB.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure EPUB save options.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,
            // Split the output at heading paragraphs to preserve chapter hierarchy.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Include headings up to level 3 in the navigation map (TOC).
            NavigationMapLevel = 3,
            // Export document properties (optional, but keeps metadata).
            ExportDocumentProperties = true
        };

        // Save as EPUB.
        pdfDoc.Save(epubPath, epubOptions);

        // Verify EPUB creation.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("EPUB file was not created.");

        // Indicate success.
        Console.WriteLine("PDF successfully converted to EPUB.");
    }
}
