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
        builder.Writeln("This is the header.");

        // Add a footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("This is the footer.");

        // Add a body paragraph.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the main body text of the document.");

        // Remove all headers and footers from the document.
        foreach (Section section in doc.Sections)
        {
            section.HeadersFooters.Clear();
        }

        // Save the document as plain text. No headers or footers will be present.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        doc.Save("Exported.txt", saveOptions);
    }
}
