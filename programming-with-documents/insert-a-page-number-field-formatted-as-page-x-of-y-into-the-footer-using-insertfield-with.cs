using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content spanning multiple pages.
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Move the builder to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert the formatted page number field: "Page X of Y".
        builder.Write("Page ");
        // PAGE field with MERGEFORMAT switch.
        builder.InsertField("PAGE  \\* MERGEFORMAT");
        builder.Write(" of ");
        // NUMPAGES field with MERGEFORMAT switch.
        builder.InsertField("NUMPAGES  \\* MERGEFORMAT");

        // Save the document to disk.
        const string outputFile = "PageNumberFooter.docx";
        doc.Save(outputFile);
    }
}
