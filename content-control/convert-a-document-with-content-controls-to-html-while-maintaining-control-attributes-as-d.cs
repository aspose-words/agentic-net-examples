using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add introductory paragraph.
        builder.Writeln("Document with content controls:");

        // Create an inline plain‑text content control.
        var inlineSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        inlineSdt.RemoveAllChildren();
        inlineSdt.AppendChild(new Run(doc, "Contoso"));
        // Append the inline SDT to the current paragraph.
        builder.CurrentParagraph.AppendChild(inlineSdt);
        builder.Writeln(); // move to next line.

        // Create a block‑level rich‑text content control.
        var blockSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "SectionContent",
            Tag = "section-content"
        };
        var blockParagraph = new Paragraph(doc);
        blockParagraph.AppendChild(new Run(doc, "This is block content control text."));
        blockSdt.AppendChild(blockParagraph);
        // Append the block SDT to the document body.
        doc.FirstSection.Body.AppendChild(blockSdt);

        // Save the source DOCX (optional, for inspection).
        doc.Save("sample.docx");

        // Prepare HTML save options to export content controls as data‑attributes.
        var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Set ExportContentControlDataAttributes if the property exists (Aspose.Words 22.9+).
        var dataAttrProp = typeof(HtmlSaveOptions).GetProperty("ExportContentControlDataAttributes");
        if (dataAttrProp != null && dataAttrProp.CanWrite)
        {
            dataAttrProp.SetValue(htmlOptions, true);
        }

        // Export the document to HTML.
        doc.Save("output.html", htmlOptions);

        // Extract content‑control metadata and write it to JSON.
        var controls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Type = sdt.SdtType.ToString(),
                Text = sdt.GetText().Trim()
            })
            .ToList();

        var json = JsonConvert.SerializeObject(controls, Formatting.Indented);
        File.WriteAllText("controls.json", json);

        Console.WriteLine("Conversion completed. Files created: sample.docx, output.html, controls.json");
    }
}
