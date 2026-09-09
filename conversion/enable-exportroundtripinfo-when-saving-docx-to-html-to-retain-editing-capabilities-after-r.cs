using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for round‑trip testing.");

        // Add header text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header text");

        // Add footer text.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer text");

        const string docxPath = "sample.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document loadedDoc = new Document(docxPath);

        // Configure HtmlSaveOptions to export round‑trip information.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportRoundtripInformation = true
        };

        const string htmlPath = "sample.html";
        loadedDoc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // Verify that round‑trip CSS information is present in the HTML.
        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("-aw-headerfooter-type"))
            throw new InvalidOperationException("Round‑trip information was not exported to HTML.");

        // Load the HTML back into a Document to ensure it can be round‑tripped.
        Document roundTripDoc = new Document(htmlPath);
        // Verify that the header and footer text are still present after loading.
        string roundTripText = roundTripDoc.GetText();
        if (!roundTripText.Contains("Header text") || !roundTripText.Contains("Footer text"))
            throw new InvalidOperationException("Header or footer information was lost during round‑trip.");

        // Cleanup temporary files (optional).
        File.Delete(docxPath);
        File.Delete(htmlPath);
    }
}
