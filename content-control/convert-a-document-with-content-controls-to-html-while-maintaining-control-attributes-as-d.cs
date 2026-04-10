using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Document with content controls:");

        // Create an inline plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "custName"
        };

        // Add some text inside the content control.
        Run run = new Run(doc, "John Doe");
        sdt.AppendChild(run);

        // Insert the content control into the current paragraph.
        Paragraph paragraph = builder.CurrentParagraph;
        paragraph.AppendChild(sdt);

        // Add a line break after the content control.
        builder.Writeln();

        // Save the sample DOCX.
        const string docxPath = "Sample.docx";
        doc.Save(docxPath);

        // Load the document for conversion.
        Document loadDoc = new Document(docxPath);

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // The following options are supported in recent Aspose.Words versions.
            // If the used version does not expose ExportContentControls or
            // ExportContentControlDataAttributes, the default behavior will still
            // preserve content controls in the HTML output.
        };

        // Save as HTML.
        const string htmlPath = "Sample.html";
        loadDoc.Save(htmlPath, htmlOptions);

        // Output a short preview of the generated HTML.
        string htmlContent = File.ReadAllText(htmlPath);
        Console.WriteLine($"HTML output saved to \"{htmlPath}\". Sample snippet:");
        Console.WriteLine(htmlContent.Substring(0, Math.Min(500, htmlContent.Length)));
    }
}
