using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different first page header/footer.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Add content to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First Page Header");

        // Add content to the primary header (used for all other pages).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Primary Header");

        // Add enough body text to generate multiple pages.
        builder.MoveToDocumentEnd();
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Save the document.
        const string outputPath = "FirstPageHeader.docx";
        doc.Save(outputPath);
    }
}
