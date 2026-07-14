using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Sample Header Text");

        // Add a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Sample Footer Text");

        // Save the document (optional, demonstrates persistence).
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "HeaderFooterSample.docx");
        doc.Save(docPath);

        // Access the header and footer objects.
        HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

        // Extract plain text from their ranges.
        string headerText = header?.Range?.Text?.Trim() ?? string.Empty;
        string footerText = footer?.Range?.Text?.Trim() ?? string.Empty;

        // Output the extracted texts.
        Console.WriteLine("Header text: " + headerText);
        Console.WriteLine("Footer text: " + footerText);
    }
}
