using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // ---------- Create a primary header ----------
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        // Add a paragraph with the header text.
        header.AppendParagraph("Sample Header Text");

        // ---------- Create a primary footer ----------
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        // Add a paragraph with the footer text.
        footer.AppendParagraph("Sample Footer Text");

        // ---------- Extract plain text from header and footer ----------
        // The Range of a HeaderFooter represents the content of that story.
        string headerText = header.Range.Text.Trim();
        string footerText = footer.Range.Text.Trim();

        // Output the extracted text (could be used for indexing).
        Console.WriteLine("Header: " + headerText);
        Console.WriteLine("Footer: " + footerText);

        // ---------- Save the document ----------
        doc.Save("HeaderFooterSample.docx");
    }
}
