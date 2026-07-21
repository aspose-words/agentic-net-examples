using System;
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

        // Move the builder to the even-page footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert the PAGE field that will display the page number.
        builder.Write("Page ");
        builder.InsertField("PAGE", "");

        // Add a few pages of content to see the footer on even pages.
        builder.MoveToSection(0);
        builder.Writeln("First page (odd).");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page (even).");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page (odd).");

        // Set the page number style to uppercase Roman numerals.
        doc.FirstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Update fields so the PAGE field shows the correct values.
        doc.UpdateFields();

        // Save the document.
        doc.Save("EvenFooterRoman.docx");
    }
}
