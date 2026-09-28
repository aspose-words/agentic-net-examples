using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Enable different first page header/footer.
        doc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First Page Header");

        // Primary (default) header for other pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Primary Header");

        // Add body content.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page.");

        // Save the document.
        string outputPath = "FirstPageHeader.docx";
        doc.Save(outputPath);

        // Verify that the document was saved and can be reopened.
        Document loadedDoc = new Document(outputPath);
        bool isDifferentFirstPage = loadedDoc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter;

        // Output verification result (no user interaction required).
        Console.WriteLine($"Document saved. DifferentFirstPageHeaderFooter = {isDifferentFirstPage}");
    }
}
