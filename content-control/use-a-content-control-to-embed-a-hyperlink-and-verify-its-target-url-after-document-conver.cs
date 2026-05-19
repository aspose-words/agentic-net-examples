using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Fields;
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
            Title = "HyperlinkControl",
            Tag = "hyperlink-sdt"
        };

        // Add a paragraph that will host the hyperlink inside the content control.
        Paragraph paragraph = new Paragraph(doc);
        sdt.AppendChild(paragraph);
        doc.FirstSection.Body.AppendChild(sdt);

        // Move the builder to the newly created paragraph and insert a hyperlink field.
        builder.MoveTo(paragraph);
        builder.Font.Color = System.Drawing.Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
        builder.Font.ClearFormatting();

        // Save the document in DOCX format.
        const string docxPath = "output.docx";
        doc.Save(docxPath);

        // Convert the document to HTML.
        const string htmlPath = "output.html";
        doc.Save(htmlPath, SaveFormat.Html);

        // Verify that the hyperlink target URL is present in the generated HTML.
        string htmlContent = File.ReadAllText(htmlPath);
        bool urlFound = htmlContent.Contains("https://www.aspose.com");

        // Serialize verification result to JSON using Newtonsoft.Json.
        var verificationResult = new
        {
            Url = "https://www.aspose.com",
            Found = urlFound
        };
        string json = JsonConvert.SerializeObject(verificationResult, Formatting.Indented);
        File.WriteAllText("verification.json", json);

        // Output the verification result to the console.
        Console.WriteLine(urlFound
            ? "Hyperlink URL verified successfully."
            : "Hyperlink URL not found in the converted document.");
    }
}
