using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class ExportContentControlsToXps
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -------------------------
        // Insert an inline plain‑text content control.
        // -------------------------
        // The first paragraph already exists in a newly created document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PlainTextControl",
            Tag = "plain-text"
        };
        plainTextSdt.RemoveAllChildren();                     // Ensure the SDT is empty.
        plainTextSdt.AppendChild(new Run(doc, "Inline plain‑text content"));
        firstParagraph.AppendChild(plainTextSdt);

        // -------------------------
        // Insert a block‑level rich‑text content control.
        // -------------------------
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "RichTextControl",
            Tag = "rich-text"
        };
        // A block‑level SDT must contain at least one block element (e.g., a paragraph).
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "Block rich‑text content"));
        richTextSdt.AppendChild(richParagraph);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the intermediate DOCX file (optional, helps verify the controls).
        doc.Save("ContentControls.docx");

        // -------------------------
        // Export the document to XPS while preserving content control boundaries.
        // -------------------------
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        // Example: limit outline to heading levels 2 (not required for control preservation).
        xpsOptions.OutlineOptions.HeadingsOutlineLevels = 2;

        doc.Save("ContentControls.xps", xpsOptions);
    }
}
