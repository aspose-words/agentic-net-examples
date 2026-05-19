using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Move the builder cursor to the even-page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Insert a PAGE field that will display the page number.
        // This overload inserts the field code without an initial result.
        builder.InsertField("PAGE", "");

        // Return to the main body of the document.
        builder.MoveToSection(0);

        // Add some content to generate multiple pages.
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Set the page number style to uppercase Roman numerals for the whole section.
        doc.FirstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Update all fields so the PAGE field shows the correct values.
        doc.UpdateFields();

        // Save the document to a file.
        doc.Save("EvenFooterRoman.docx");
    }
}
