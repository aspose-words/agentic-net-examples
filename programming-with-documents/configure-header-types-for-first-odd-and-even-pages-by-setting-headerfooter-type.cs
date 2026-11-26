using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace HeaderTypesExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different headers for first page and odd/even pages.
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // First page header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Writeln("First Page Header");

            // Odd page header (primary header).
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Odd Page Header");

            // Even page header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Writeln("Even Page Header");

            // Add body content to generate multiple pages.
            builder.MoveToDocumentEnd();
            builder.Writeln("This is page 1 content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 2 content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 3 content.");

            // Save the document.
            string outputPath = "HeaderTypes.docx";
            doc.Save(outputPath);
        }
    }
}
