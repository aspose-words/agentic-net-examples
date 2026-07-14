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

        // Use DocumentBuilder for convenient content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a block-level RichText content control.
        StructuredDocumentTag blockSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "BlockRichText",
            Tag = "block-rich"
        };
        // The block SDT must contain at least one paragraph.
        Paragraph blockParagraph = new Paragraph(doc);
        blockParagraph.AppendChild(new Run(doc, "This is text inside a block-level RichText content control."));
        blockSdt.AppendChild(blockParagraph);
        // Append the block SDT to the document body.
        doc.FirstSection.Body.AppendChild(blockSdt);

        // Insert an inline PlainText content control.
        StructuredDocumentTag inlineSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "InlinePlainText",
            Tag = "inline-plain"
        };
        inlineSdt.RemoveAllChildren();
        inlineSdt.AppendChild(new Run(doc, "Inline plain text"));
        // Add the inline SDT to the current paragraph.
        builder.Writeln(); // Ensure we are on a new paragraph.
        Paragraph inlineParagraph = doc.FirstSection.Body.LastParagraph;
        inlineParagraph.AppendChild(inlineSdt);

        // Insert a checkbox content control.
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AgreeCheckBox",
            Tag = "agree-checkbox",
            Checked = false
        };
        // Add some descriptive text before the checkbox.
        builder.Writeln(); // New paragraph for clarity.
        Paragraph checkBoxParagraph = doc.FirstSection.Body.LastParagraph;
        checkBoxParagraph.AppendChild(new Run(doc, "I agree to the terms: "));
        checkBoxParagraph.AppendChild(checkBoxSdt);

        // Optional: Output information about the content controls to the console.
        foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
        {
            Console.WriteLine($"Title: {sdt.Title}, Tag: {sdt.Tag}, Type: {sdt.SdtType}");
            Console.WriteLine($"XML (minimal): {sdt.WordOpenXMLMinimal}");
        }

        // Prepare XPS save options. No special options are required to preserve control boundaries.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the document as XPS in the working directory.
        string xpsPath = Path.Combine(Directory.GetCurrentDirectory(), "ContentControls.xps");
        doc.Save(xpsPath, xpsOptions);

        // Indicate completion.
        Console.WriteLine($"Document saved to XPS: {xpsPath}");
    }
}
