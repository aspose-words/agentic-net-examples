using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // NOTE: The property to enable different odd/even footers may not be available
        // in some older Aspose.Words versions. Even without setting it, we can still
        // insert content into the even‑page footer; Word will display it on even pages
        // when the setting is enabled in the resulting file.

        // Add some text to generate multiple pages.
        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"This is page {i + 1}");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Move the builder to the even‑page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        // Insert a PAGE field formatted with Roman numerals.
        builder.InsertField("PAGE  \\* ROMAN");

        // Save the document.
        const string outputPath = "EvenPageFooterRoman.docx";
        doc.Save(outputPath);
    }
}
