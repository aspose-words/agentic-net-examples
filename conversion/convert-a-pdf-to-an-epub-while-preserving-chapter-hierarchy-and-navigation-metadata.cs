using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string pdfPath = "sample.pdf";
        const string epubPath = "output.epub";

        // -------------------------------------------------
        // 1. Create a sample document with heading hierarchy
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Section 1.1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details of section 1.1.");

        // Subsection 1.1.1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("More details.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Save the document as PDF (input for conversion)
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 2. Load the PDF and convert it to EPUB
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Configure EPUB save options to preserve hierarchy and navigation metadata
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            // Split the EPUB into parts at heading paragraphs (preserves chapters)
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Export built‑in and custom document properties
            ExportDocumentProperties = true,
            // Include headings up to level 3 in the navigation map (TOC)
            NavigationMapLevel = 3,
            // Use UTF‑8 encoding
            Encoding = Encoding.UTF8
        };

        // Perform the conversion
        pdfDoc.Save(epubPath, epubOptions);

        // -------------------------------------------------
        // 3. Validate the output
        // -------------------------------------------------
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("EPUB file was not created.");

        Console.WriteLine("PDF successfully converted to EPUB: " + Path.GetFullPath(epubPath));
    }
}
