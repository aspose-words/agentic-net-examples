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
        builder.Writeln("Header: Sample Header Text");

        // Return to the main body.
        builder.MoveToDocumentEnd();

        // Add a body paragraph.
        builder.Writeln("Body: This is the main content of the document.");

        // Add a footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer: Sample Footer Text");

        // Configure TXT save options.
        // In recent versions of Aspose.Words the ExportHeadersFooters property is available.
        // If the property is not present in the referenced version, headers and footers are
        // included by default when saving to TXT, so we can proceed without setting it.
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Save the document as a TXT file.
        string outputFile = "DocumentWithHeadersFooters.txt";
        doc.Save(outputFile, saveOptions);

        Console.WriteLine($"Document saved to {outputFile}");
    }
}
