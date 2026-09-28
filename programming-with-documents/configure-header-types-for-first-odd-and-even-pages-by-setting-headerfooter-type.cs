using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers for the first page and for odd/even pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // ----- First page header -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First Page Header");

        // ----- Odd pages header (primary) -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Odd Page Header");

        // ----- Even pages header -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Writeln("Even Page Header");

        // Save the document to a file.
        string outputPath = "HeaderTypes.docx";
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
        }
    }
}
