using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for convenient paragraph handling.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with content controls:");

        // ----- Inline plain‑text content control -----
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "John Doe"));
        // Insert the inline SDT at the current cursor position.
        builder.InsertNode(plainTextSdt);

        // ----- Inline checkbox content control -----
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AgreeTerms",
            Tag = "agree-terms",
            Checked = true
        };
        builder.InsertNode(checkBoxSdt);

        // Add a blank line for readability.
        builder.Writeln();

        // ----- Block‑level rich‑text content control -----
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Notes",
            Tag = "notes"
        };
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "This is a rich text content control."));
        richTextSdt.AppendChild(richParagraph);
        // Append the block‑level SDT directly to the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Optional: display each content control's minimal WordOpenXML to verify boundaries.
        foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
        {
            Console.WriteLine($"SDT Title: {sdt.Title}");
            Console.WriteLine(sdt.WordOpenXMLMinimal);
        }

        // Save the document to XPS format, preserving content control boundaries.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        doc.Save("ContentControls.xps", xpsOptions);
    }
}
