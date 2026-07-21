using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one section.
        doc.EnsureMinimum();

        // Set narrow margins for better layout on narrow pages.
        doc.FirstSection.PageSetup.Margins = Margins.Narrow;

        // Insert a paragraph with sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Apply justified alignment.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        // Enable word wrap (wrap by whole words). This is true by default, but set explicitly.
        builder.ParagraphFormat.WordWrap = true;

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "JustifiedParagraph.docx");
        doc.Save(outputPath);
    }
}
