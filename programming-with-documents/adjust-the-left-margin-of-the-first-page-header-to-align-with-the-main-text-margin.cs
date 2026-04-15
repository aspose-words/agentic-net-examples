using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output directory and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "FirstPageHeaderAligned.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the main body left margin (e.g., 1 inch).
        double leftMargin = ConvertUtil.InchToPoint(1.0);
        builder.PageSetup.LeftMargin = leftMargin;

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move to the first‑page header and add some text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First page header");

        // Align the header's left margin with the main text margin
        // by setting the left indent of the header paragraph.
        HeaderFooter firstHeader = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderFirst];
        if (firstHeader?.FirstParagraph != null)
        {
            firstHeader.FirstParagraph.ParagraphFormat.LeftIndent = leftMargin;
        }

        // Return to the main body and add regular content.
        builder.MoveToSection(0);
        builder.Writeln("Main document content.");

        // Save the document.
        doc.Save(outputPath);
    }
}
