using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF input and the resulting EPUB output.
        const string pdfPath = "sample_input.pdf";
        const string epubPath = "converted_output.epub";

        // -----------------------------------------------------------------
        // 1. Create a sample PDF document with a simple chapter hierarchy.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the introduction chapter.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1: Overview");
        builder.Writeln("Some overview text.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.2: Deep Dive");
        builder.Writeln("Detailed information goes here.");

        // Save the document as PDF.
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // ---------------------------------------------------------------
        // 2. Load the PDF and convert it to EPUB while preserving hierarchy.
        // ---------------------------------------------------------------
        Document pdfDocument = new Document(pdfPath);

        // Configure EPUB save options.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            // Split the output at heading paragraphs to keep chapter structure.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Include up to three heading levels in the navigation map (TOC).
            NavigationMapLevel = 3,
            // Export built‑in and custom document properties.
            ExportDocumentProperties = true
        };

        // Save as EPUB.
        pdfDocument.Save(epubPath, epubOptions);

        // Verify that the EPUB was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("EPUB conversion failed; output file not found.");

        // Clean up temporary PDF if desired.
        // File.Delete(pdfPath);
    }
}
