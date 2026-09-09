using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the page margins (example: 1 inch left margin).
        builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1.0);
        builder.PageSetup.RightMargin = ConvertUtil.InchToPoint(1.0);
        builder.PageSetup.TopMargin = ConvertUtil.InchToPoint(1.0);
        builder.PageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0);

        // Enable a different header/footer for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move to the first‑page header and add some text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First page header");

        // Align the left margin of the first‑page header with the main text margin.
        HeaderFooter firstHeader = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderFirst];
        if (firstHeader != null && firstHeader.FirstParagraph != null)
        {
            // Set the left indent of the header paragraph to the section's left margin.
            firstHeader.FirstParagraph.ParagraphFormat.LeftIndent = doc.FirstSection.PageSetup.LeftMargin;
        }

        // Return to the main document body and add some content.
        builder.MoveToSection(0);
        builder.Writeln("Body text aligned with the same left margin.");

        // Define an output path for the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FirstPageHeaderAligned.docx");
        // Ensure the directory exists.
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
