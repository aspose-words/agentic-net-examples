using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "LinkControl",
            Tag = "link-control"
        };

        // Add a paragraph that will hold the hyperlink.
        Paragraph innerParagraph = new Paragraph(doc);
        sdt.AppendChild(innerParagraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Move the builder to the paragraph inside the content control.
        builder.MoveTo(innerParagraph);

        // Insert a hyperlink into the paragraph.
        const string displayText = "Aspose";
        const string targetUrl = "https://www.aspose.com";
        builder.InsertHyperlink(displayText, targetUrl, false);

        // Save the document in DOCX and HTML formats.
        const string docxPath = "output.docx";
        const string htmlPath = "output.html";
        doc.Save(docxPath, SaveFormat.Docx);
        doc.Save(htmlPath, SaveFormat.Html);

        // Verify that the HTML output contains the expected hyperlink URL.
        string htmlContent = File.ReadAllText(htmlPath);
        bool urlFound = htmlContent.Contains(targetUrl, StringComparison.OrdinalIgnoreCase);

        // Prepare verification result.
        var verificationResult = new
        {
            Document = docxPath,
            HtmlFile = htmlPath,
            TargetUrl = targetUrl,
            UrlFoundInHtml = urlFound,
            Message = urlFound
                ? "Hyperlink URL verified successfully."
                : "Hyperlink URL not found in HTML output."
        };

        // Serialize result to JSON file.
        string jsonPath = "verification.json";
        File.WriteAllText(jsonPath,
            JsonConvert.SerializeObject(verificationResult, Formatting.Indented));

        // Output result to console.
        Console.WriteLine(JsonConvert.SerializeObject(verificationResult, Formatting.Indented));
    }
}
