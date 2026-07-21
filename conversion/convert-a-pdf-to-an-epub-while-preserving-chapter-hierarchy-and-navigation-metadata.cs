using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        string pdfPath = "sample.pdf";
        string epubPath = "output.epub";

        // -----------------------------------------------------------------
        // 1. Create a sample PDF with a simple chapter hierarchy.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Section 1.1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Content of section 1.1.");

        // Section 1.2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2");
        builder.Writeln("Content of section 1.2.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Section 2.1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");
        builder.Writeln("Content of section 2.1.");

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to EPUB while preserving hierarchy.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        HtmlSaveOptions epubOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,
            Encoding = Encoding.UTF8,
            // Split the EPUB at heading paragraphs to keep chapter structure.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Export document properties (optional, but useful for metadata).
            ExportDocumentProperties = true,
            // Include headings up to level 3 in the navigation map.
            NavigationMapLevel = 3
        };

        pdfDoc.Save(epubPath, epubOptions);

        // Verify EPUB creation.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("EPUB conversion failed; output file not found.");

        // Cleanup temporary PDF if desired.
        // File.Delete(pdfPath);
    }
}
