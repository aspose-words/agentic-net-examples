using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // ---------- Insert an inline plain‑text content control ----------
        Paragraph inlineParagraph = doc.FirstSection.Body.FirstParagraph;
        StructuredDocumentTag inlineSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        inlineSdt.RemoveAllChildren();
        inlineSdt.AppendChild(new Run(doc, "Contoso"));
        inlineParagraph.AppendChild(inlineSdt);

        // ---------- Insert a block‑level rich‑text content control ----------
        StructuredDocumentTag blockSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "SectionContent",
            Tag = "section-content"
        };
        Paragraph blockParagraph = new Paragraph(doc);
        blockParagraph.AppendChild(new Run(doc, "This is block‑level content control text."));
        blockSdt.AppendChild(blockParagraph);
        doc.FirstSection.Body.AppendChild(blockSdt);

        // ---------- Insert an inline drop‑down list content control ----------
        StructuredDocumentTag dropdownSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Options",
            Tag = "options"
        };
        dropdownSdt.ListItems.Add(new SdtListItem("Option A", "A"));
        dropdownSdt.ListItems.Add(new SdtListItem("Option B", "B"));
        // Set default displayed text.
        dropdownSdt.AppendChild(new Run(doc, "Option A"));
        inlineParagraph.AppendChild(dropdownSdt);

        // Save the sample DOCX to the working directory.
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        doc.Save(docxPath);

        // Load the document back (simulating processing an existing file).
        Document loadedDoc = new Document(docxPath);

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // The property ExportContentControlsAsDataAttributes is not available in the
            // current Aspose.Words version, so we rely on the default behavior.
            ExportImagesAsBase64 = true // Embed images directly for a self‑contained HTML file.
        };

        // Save the document as HTML.
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.html");
        loadedDoc.Save(htmlPath, htmlOptions);
    }
}
