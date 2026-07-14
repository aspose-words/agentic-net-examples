using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Create the even‑page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Page ");
        // Insert a PAGE field; the result will be updated when the document is saved or fields are updated.
        builder.InsertField("PAGE", "");

        // Return to the main body of the document.
        builder.MoveToSection(0);

        // Set the page numbering style to uppercase Roman numerals.
        doc.Sections[0].PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Add some content to generate multiple pages.
        builder.Writeln("First page (odd).");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page (even).");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page (odd).");

        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "EvenPageFooterRoman.docx");
        doc.Save(outputPath);
    }
}
