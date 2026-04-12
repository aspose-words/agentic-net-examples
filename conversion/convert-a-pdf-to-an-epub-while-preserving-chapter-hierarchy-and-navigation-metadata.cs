using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.pdf");
        string epubPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.epub");

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document with heading styles.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDoc);

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
        pdfDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the PDF file.", pdfPath);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to EPUB.
        // -----------------------------------------------------------------
        Document loadedPdf = new Document(pdfPath);

        // Configure EPUB save options.
        HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            // Split the output at heading paragraphs to preserve chapter hierarchy.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Include headings up to level 3 in the navigation map (TOC).
            NavigationMapLevel = 3,
            // Export document properties (optional, but often useful).
            ExportDocumentProperties = true,
            // Use UTF-8 encoding for the EPUB content.
            Encoding = Encoding.UTF8
        };

        // Save as EPUB using the configured options.
        loadedPdf.Save(epubPath, epubSaveOptions);

        // -----------------------------------------------------------------
        // Step 3: Validate the EPUB output.
        // -----------------------------------------------------------------
        if (!File.Exists(epubPath))
            throw new FileNotFoundException("EPUB conversion failed; output file not found.", epubPath);

        // Optionally, report success.
        Console.WriteLine("PDF successfully converted to EPUB:");
        Console.WriteLine($"PDF path : {pdfPath}");
        Console.WriteLine($"EPUB path: {epubPath}");
    }
}
