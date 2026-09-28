using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header text");

        // Add body content.
        builder.MoveToDocumentEnd();
        builder.Writeln("Body paragraph.");

        // Add a footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer text");

        // Remove all headers and footers before saving to TXT.
        foreach (Section section in doc.Sections)
        {
            // Clear the collection of headers and footers for this section.
            section.HeadersFooters.Clear();
        }

        // Configure TXT save options (no special options needed now).
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Export the document to a TXT file without headers and footers.
        doc.Save("Output.txt", saveOptions);
    }
}
