using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class HeaderFooterIndexer
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Indexable Header Text");

        // Add a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Indexable Footer Text");

        // Add some body content (optional).
        builder.MoveToDocumentEnd();
        builder.Writeln("Body content does not affect header/footer extraction.");

        // Save the document locally (optional, demonstrates persistence).
        string docPath = "HeaderFooterSample.docx";
        doc.Save(docPath, SaveFormat.Docx);

        // Retrieve the header and footer objects.
        HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

        // Extract plain text from the header and footer using their Range objects.
        string headerText = header?.Range?.Text?.Trim() ?? string.Empty;
        string footerText = footer?.Range?.Text?.Trim() ?? string.Empty;

        // Combine the extracted texts for indexing purposes.
        string indexableText = $"{headerText} {footerText}".Trim();

        // Output the extracted texts.
        Console.WriteLine("Header Text: " + headerText);
        Console.WriteLine("Footer Text: " + footerText);
        Console.WriteLine("Combined Indexable Text: " + indexableText);
    }
}
