using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with a default section.
        Document doc = new Document();
        Section section = doc.FirstSection;

        // Add a primary header with sample text.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        Paragraph headerPara = new Paragraph(doc);
        headerPara.AppendChild(new Run(doc, "Header Text for Indexing"));
        header.AppendChild(headerPara);
        section.HeadersFooters.Add(header);

        // Add a primary footer with sample text.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        Paragraph footerPara = new Paragraph(doc);
        footerPara.AppendChild(new Run(doc, "Footer Text for Indexing"));
        footer.AppendChild(footerPara);
        section.HeadersFooters.Add(footer);

        // Save the document to a local file.
        const string docPath = "Sample.docx";
        doc.Save(docPath);

        // Load the document back (demonstrates loading workflow).
        Document loadedDoc = new Document(docPath);

        // Extract plain text from the header using its Range.
        HeaderFooter loadedHeader = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        string headerText = loadedHeader?.Range?.Text?.Trim() ?? string.Empty;

        // Extract plain text from the footer using its Range.
        HeaderFooter loadedFooter = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        string footerText = loadedFooter?.Range?.Text?.Trim() ?? string.Empty;

        // Output the extracted texts.
        Console.WriteLine("Extracted Header Text: " + headerText);
        Console.WriteLine("Extracted Footer Text: " + footerText);
    }
}
