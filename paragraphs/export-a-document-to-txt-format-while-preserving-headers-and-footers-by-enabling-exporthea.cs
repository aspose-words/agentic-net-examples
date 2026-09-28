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

        // Add content to the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header: Document Title");

        // Return to the main body and add some paragraphs.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the first paragraph of the document body.");
        builder.Writeln("This is the second paragraph of the document body.");

        // Add content to the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer: Confidential");

        // Configure TXT save options.
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // In newer versions of Aspose.Words the ExportHeadersFooters property can be set to true:
        // saveOptions.ExportHeadersFooters = true;
        // If the property is not available in the referenced version, headers and footers are
        // included by default when saving to TXT.

        // Save the document as a TXT file.
        doc.Save("ExportedDocument.txt", saveOptions);
    }
}
