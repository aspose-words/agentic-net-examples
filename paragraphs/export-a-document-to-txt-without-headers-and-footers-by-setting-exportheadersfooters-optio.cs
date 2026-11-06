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

        // Add a header to the document.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("This is the header.");

        // Add a body paragraph.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the main body paragraph.");

        // Add a footer to the document.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("This is the footer.");

        // Remove all headers and footers from the document before saving.
        foreach (Section section in doc.Sections)
        {
            section.HeadersFooters.Clear();
        }

        // Configure TXT save options (no special options needed now).
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Save the document as a TXT file without headers and footers.
        doc.Save("output.txt", saveOptions);
    }
}
