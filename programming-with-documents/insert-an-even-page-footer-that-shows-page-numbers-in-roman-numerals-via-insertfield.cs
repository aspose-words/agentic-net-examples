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

        // Add enough content to generate several pages.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Set the page number style for the section to uppercase Roman numerals.
        doc.FirstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Move the builder to the even‑page footer and insert a PAGE field.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        builder.Write("Page ");
        builder.InsertField("PAGE", "");

        // Save the resulting document.
        doc.Save("EvenPageFooter.docx");
    }
}
