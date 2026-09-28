using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class ContentControlSummary
{
    // Simple DTO for JSON serialization.
    private class SummaryItem
    {
        public string Title { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document that contains a variety of content controls.
        // -----------------------------------------------------------------
        Document doc = new Document();
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Plain text inline content control.
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        plainTextSdt.Title = "PlainText";
        plainTextSdt.Tag = "plain";
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "Sample plain text"));
        firstParagraph.AppendChild(plainTextSdt);

        // Rich text block‑level content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richTextSdt.Title = "RichText";
        richTextSdt.Tag = "rich";
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "Sample rich text block"));
        richTextSdt.AppendChild(richParagraph);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Drop‑down list inline content control.
        StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline);
        dropDownSdt.Title = "DropDown";
        dropDownSdt.Tag = "dropdown";
        dropDownSdt.ListItems.Add(new SdtListItem("Option 1", "1"));
        dropDownSdt.ListItems.Add(new SdtListItem("Option 2", "2"));
        firstParagraph.AppendChild(dropDownSdt);

        // Checkbox inline content control.
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
        checkBoxSdt.Title = "CheckBox";
        checkBoxSdt.Tag = "checkbox";
        checkBoxSdt.Checked = true;
        firstParagraph.AppendChild(checkBoxSdt);

        // Date placeholder – using a plain‑text SDT because the DateTime type is not available in this version.
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        dateSdt.Title = "DateControl";
        dateSdt.Tag = "date";
        dateSdt.RemoveAllChildren();
        dateSdt.AppendChild(new Run(doc, DateTime.Now.ToShortDateString()));
        firstParagraph.AppendChild(dateSdt);

        // Save the sample document.
        const string samplePath = "sample.docx";
        doc.Save(samplePath);

        // -----------------------------------------------------------------
        // 2. Load the document and build a summary of all content controls.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(samplePath);
        List<SummaryItem> summary = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(sdt => new SummaryItem
            {
                Title = sdt.Title ?? string.Empty,
                Tag = sdt.Tag ?? string.Empty,
                Type = sdt.SdtType.ToString(),
                Text = sdt.GetText().Trim()
            })
            .ToList();

        // -----------------------------------------------------------------
        // 3. Serialize the summary to JSON and write to a file.
        // -----------------------------------------------------------------
        string json = JsonConvert.SerializeObject(summary, Formatting.Indented);
        const string jsonPath = "content-controls-summary.json";
        File.WriteAllText(jsonPath, json);

        // -----------------------------------------------------------------
        // 4. Write a simple console report.
        // -----------------------------------------------------------------
        Console.WriteLine("Content Control Summary:");
        foreach (var item in summary)
        {
            Console.WriteLine($"Title: {item.Title}, Tag: {item.Tag}, Type: {item.Type}, Text: {item.Text}");
        }
        Console.WriteLine($"Summary written to {jsonPath}");
    }
}
