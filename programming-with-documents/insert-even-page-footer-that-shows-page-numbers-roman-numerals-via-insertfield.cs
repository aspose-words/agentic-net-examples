using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Move the builder's cursor to the even‑page footer of the current section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);

        // Write static text before the page number (optional).
        builder.Write("Page ");

        // Insert a PAGE field that will display the current page number.
        // The second argument is an empty string because we do not provide a placeholder value.
        builder.InsertField("PAGE", "");

        // Set the page number format for the whole document to uppercase Roman numerals.
        // This affects all PAGE fields, including the one we just inserted.
        doc.FirstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

        // Save the document to disk.
        doc.Save("EvenFooterRoman.docx");
    }
}
