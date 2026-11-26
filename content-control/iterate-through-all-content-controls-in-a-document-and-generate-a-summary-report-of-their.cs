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
        StructuredDocumentTag plainText = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        plainText.Title = "CustomerName";
        plainText.Tag = "customer-name";
        plainText.RemoveAllChildren();
        plainText.AppendChild(new Run(doc, "Contoso"));
        builder.Writeln("Plain text control:");
        builder.CurrentParagraph.AppendChild(plainText);

        // Rich text block content control.
        StructuredDocumentTag richText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richText.Title = "SectionContent";
        richText.Tag = "section-content";
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "This is a rich text block content control."));
        richText.AppendChild(richParagraph);
        builder.Writeln();
        doc.FirstSection.Body.AppendChild(richText);

        // Checkbox inline content control.
        StructuredDocumentTag checkbox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
        checkbox.Title = "AgreeTerms";
        checkbox.Tag = "agree-terms";
        checkbox.Checked = true;
        builder.Writeln();
        builder.Writeln("Checkbox control:");
        builder.CurrentParagraph.AppendChild(checkbox);

        // Drop-down list inline content control.
        StructuredDocumentTag dropdown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline);
        dropdown.Title = "Country";
        dropdown.Tag = "country";
        dropdown.ListItems.Add(new SdtListItem("USA", "US"));
        dropdown.ListItems.Add(new SdtListItem("Canada", "CA"));
        dropdown.ListItems.Add(new SdtListItem("Mexico", "MX"));
        builder.Writeln();
        builder.Writeln("Drop-down list control:");
        builder.CurrentParagraph.AppendChild(dropdown);

        // Repeating section block content control.
        StructuredDocumentTag repeating = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeating.Title = "Items";
        repeating.Tag = "items";
        Paragraph itemParagraph = new Paragraph(doc);
        itemParagraph.AppendChild(new Run(doc, "Item 1"));
        repeating.AppendChild(itemParagraph);
        doc.FirstSection.Body.AppendChild(repeating);

        // Save the sample document.
        const string samplePath = "sample.docx";
        doc.Save(samplePath);

        // Load the document (simulating processing an existing file).
        Document loadedDoc = new Document(samplePath);

        // Gather summary information about all content controls.
        var summary = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(sdt => new
            {
                Type = sdt.SdtType.ToString(),
                Title = sdt.Title,
                Tag = sdt.Tag,
                // Additional details for specific types.
                IsChecked = sdt.SdtType == SdtType.Checkbox ? sdt.Checked : (bool?)null,
                ListItemCount = (sdt.SdtType == SdtType.DropDownList || sdt.SdtType == SdtType.ComboBox) ? sdt.ListItems.Count : (int?)null
            })
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
