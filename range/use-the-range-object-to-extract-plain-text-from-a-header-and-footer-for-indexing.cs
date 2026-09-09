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

        // Create and add a primary header with sample text.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        header.AppendParagraph("Sample Header Text");

        // Create and add a primary footer with sample text.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        footer.AppendParagraph("Sample Footer Text");

        // Add a body paragraph to ensure the document has content.
        builder.Writeln("Body paragraph.");

        // Extract plain text from the header and footer using their Range objects.
        string headerText = header.Range.Text.Trim();
        string footerText = footer.Range.Text.Trim();

        // Output the extracted texts.
        Console.WriteLine("Header text: " + headerText);
        Console.WriteLine("Footer text: " + footerText);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HeaderFooterSample.docx");
        doc.Save(outputPath);
    }
}
