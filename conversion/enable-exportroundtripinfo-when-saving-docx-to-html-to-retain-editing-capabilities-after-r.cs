using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportRoundTripExample
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document.");
        // Add a header to demonstrate round‑trip information.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Sample Header");
        // Save the document as DOCX.
        string docxPath = "sample.docx";
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(docxPath);

        // Configure HTML save options to export round‑trip information.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportRoundtripInformation = true
        };

        // Save the document as HTML.
        string htmlPath = "sample.html";
        doc.Save(htmlPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // Optional: Load the HTML back to ensure round‑trip data is present.
        Document roundTripDoc = new Document(htmlPath);
        // Verify that the header text is preserved after round‑trip.
        string headerText = roundTripDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetText().Trim();
        if (!headerText.Equals("Sample Header"))
            throw new InvalidOperationException("Round‑trip information was not preserved.");

        // Clean up temporary files (optional).
        File.Delete(docxPath);
        File.Delete(htmlPath);
    }
}
