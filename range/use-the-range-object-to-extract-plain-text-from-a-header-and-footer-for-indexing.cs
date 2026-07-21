using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body text.
        builder.Writeln("Body content.");

        // Create a primary header and add text.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        header.AppendParagraph("Header for indexing.");

        // Create a primary footer and add text.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        footer.AppendParagraph("Footer for indexing.");

        // Save the document.
        doc.Save("HeaderFooterSample.docx");

        // Extract plain text from header and footer using their Range objects.
        string headerText = header.Range.Text.Trim();
        string footerText = footer.Range.Text.Trim();

        // Combine the extracted texts for indexing purposes.
        string indexableText = $"{headerText} {footerText}";

        // Output the combined text.
        Console.WriteLine(indexableText);
    }
}
