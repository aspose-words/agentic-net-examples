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

        // Subsection 1.1.1 (Heading 3)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1");

        // Chapter 2 (Heading 1)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Section 2.1 (Heading 2)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to EPUB.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure EPUB save options.
        HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,
            Encoding = Encoding.UTF8,
            // Split the output at heading paragraphs to preserve chapter hierarchy.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Export built‑in and custom document properties.
            ExportDocumentProperties = true,
            // Define how many heading levels appear in the navigation map.
            NavigationMapLevel = 3
        };

        // Save as EPUB.
        pdfDoc.Save(epubPath, epubSaveOptions);

        // Verify that the EPUB was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("The EPUB file was not created.");

        // The example finishes without waiting for user input.
    }
}
