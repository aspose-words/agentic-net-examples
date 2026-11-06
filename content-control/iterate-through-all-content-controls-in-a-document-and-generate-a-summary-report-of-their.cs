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
        // Create a sample document with various content controls.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Plain text inline content control.
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "Contoso"));
        builder.CurrentParagraph.AppendChild(plainTextSdt);
        builder.Writeln();

        // Rich text block content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "SectionContent",
            Tag = "section-content"
        };
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "This is a rich text block content control."));
        richTextSdt.AppendChild(richParagraph);
        doc.FirstSection.Body.AppendChild(richTextSdt);
        builder.Writeln();

        // Drop‑down list inline content control.
        StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Country",
            Tag = "country"
        };
        dropDownSdt.ListItems.Add(new SdtListItem("USA", "US"));
        dropDownSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
        dropDownSdt.ListItems.Add(new SdtListItem("Mexico", "MX"));
        builder.CurrentParagraph.AppendChild(dropDownSdt);
        builder.Writeln();

        // Checkbox inline content control.
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Subscribe",
            Tag = "subscribe",
            Checked = true
        };
        builder.CurrentParagraph.AppendChild(checkBoxSdt);
        builder.Writeln();

        // Date inline content control (using default settings).
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
        {
            Title = "BirthDate",
            Tag = "birth-date"
        };
        builder.CurrentParagraph.AppendChild(dateSdt);
        builder.Writeln();

        // Repeating section block content control.
        StructuredDocumentTag repeatingSdt = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "Items",
            Tag = "items"
        };
        Paragraph itemParagraph = new Paragraph(doc);
        itemParagraph.AppendChild(new Run(doc, "Item 1"));
        repeatingSdt.AppendChild(itemParagraph);
        doc.FirstSection.Body.AppendChild(repeatingSdt);
        builder.Writeln();

        // Save the sample document.
        const string samplePath = "sample.docx";
        doc.Save(samplePath);

        // Load the document for processing.
        Document loadedDoc = new Document(samplePath);

        // Enumerate all StructuredDocumentTag nodes and collect summary information.
        List<object> summary = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Type = sdt.SdtType.ToString(),
                Appearance = sdt.Appearance.ToString()
            })
            .Cast<object>()
            .ToList();

        // Serialize the summary to JSON.
        string json = JsonConvert.SerializeObject(summary, Formatting.Indented);
        const string jsonPath = "content-controls-summary.json";
        File.WriteAllText(jsonPath, json);

        // Output the summary to the console.
        Console.WriteLine("Content Control Summary:");
        Console.WriteLine(json);
    }
}
