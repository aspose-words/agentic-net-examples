using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to header, body, and footer.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header Text");

        // Add a body paragraph.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is a body paragraph.");

        // Add a footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer Text");

        // Configure TXT save options.
        // Note: In some older versions of Aspose.Words the ExportHeadersFooters property
        // may not be available. If it exists, you can enable it as shown below.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        // Uncomment the following line if the property is supported in your version:
        // saveOptions.ExportHeadersFooters = true;

        // Save the document as a TXT file.
        doc.Save("output.txt", saveOptions);
    }
}
