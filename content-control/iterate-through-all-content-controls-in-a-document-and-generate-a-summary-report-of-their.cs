using System;
using System.Collections.Generic;
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

        // ---------- Plain text inline content control ----------
        builder.Writeln("Plain text content control:");
        // Create an inline plain‑text SDT and add it to the current paragraph.
        StructuredDocumentTag plain = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PlainText",
            Tag = "plain"
        };
        plain.AppendChild(new Run(doc, "Sample plain text"));
        builder.CurrentParagraph.AppendChild(plain);

        // ---------- Rich text block content control ----------
        builder.Writeln("Rich text content control:");
        // Create a block‑level rich‑text SDT and add a paragraph with text inside it.
        StructuredDocumentTag rich = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "RichText",
            Tag = "rich"
        };
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "Sample rich text"));
        rich.AppendChild(richParagraph);
        // Block‑level SDTs are children of the document body.
        doc.FirstSection.Body.AppendChild(rich);

        // ---------- Checkbox inline content control ----------
        builder.Writeln("Checkbox content control:");
        StructuredDocumentTag check = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "CheckBox",
            Tag = "checkbox",
            Checked = true
        };
        builder.CurrentParagraph.AppendChild(check);

        // ---------- Drop‑down list inline content control ----------
        builder.Writeln("Drop‑down list content control:");
        StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "DropDown",
            Tag = "dropdown"
        };
        dropdown.ListItems.Add(new SdtListItem("Option 1", "1"));
        dropdown.ListItems.Add(new SdtListItem("Option 2", "2"));
        dropdown.ListItems.Add(new SdtListItem("Option 3", "3"));
        builder.CurrentParagraph.AppendChild(dropdown);

        // ---------- Date picker inline content control ----------
        builder.Writeln("Date picker content control:");
        StructuredDocumentTag date = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
        {
            Title = "DatePicker",
            Tag = "date"
        };
        builder.CurrentParagraph.AppendChild(date);

        // ---------- Picture inline content control ----------
        builder.Writeln("Picture content control:");
        StructuredDocumentTag picture = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Inline)
        {
            Title = "Picture",
            Tag = "picture"
        };
        builder.CurrentParagraph.AppendChild(picture);

        // Save the sample document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document (demonstrates iteration on a separate instance).
        Document loadedDoc = new Document(docPath);

        // Collect information about each content control.
        NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        var reportItems = new List<object>();

        foreach (StructuredDocumentTag sdt in sdtNodes.OfType<StructuredDocumentTag>())
        {
            reportItems.Add(new
            {
                Title = sdt.Title ?? string.Empty,
                Tag = sdt.Tag ?? string.Empty,
                Type = sdt.SdtType.ToString(),
                Content = sdt.GetText().Trim()
            });
        }

        // Serialize the report to JSON.
        string jsonReport = JsonConvert.SerializeObject(reportItems, Formatting.Indented);
        const string jsonPath = "content-controls-report.json";
        File.WriteAllText(jsonPath, jsonReport);

        // Write a brief summary to the console.
        Console.WriteLine($"Processed {reportItems.Count} content controls.");
        Console.WriteLine($"Report saved to '{jsonPath}'.");
    }
}
