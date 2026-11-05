using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an inline plain‑text content control.
        StructuredDocumentTag plainSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        plainSdt.Title = "CustomerName";
        plainSdt.Tag = "customer-name";
        plainSdt.RemoveAllChildren();
        plainSdt.AppendChild(new Run(doc, "Contoso"));
        // Insert the inline SDT into the current paragraph.
        builder.CurrentParagraph?.AppendChild(plainSdt);
        builder.Writeln(); // start a new paragraph.

        // Add a block‑level rich‑text content control.
        StructuredDocumentTag richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richSdt.Title = "SectionContent";
        richSdt.Tag = "section-content";
        Paragraph innerParagraph = new Paragraph(doc);
        innerParagraph.AppendChild(new Run(doc, "This is block content control."));
        richSdt.AppendChild(innerParagraph);
        // Insert the block SDT into the document body.
        doc.FirstSection.Body.AppendChild(richSdt);

        // Save the source DOCX.
        doc.Save("sample.docx");

        // Load the document and convert to HTML.
        Document loadedDoc = new Document("sample.docx");
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // In newer versions this option can be set to export content control
            // attributes as data‑attributes. If the property is unavailable, the
            // default behavior still preserves the attributes in the HTML.
        };
        loadedDoc.Save("sample.html", htmlOptions);
    }
}
