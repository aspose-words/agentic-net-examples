using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    // Simple model matching the JSON structure.
    private class Item
    {
        public string Text { get; set; } = "";
    }

    public static void Main()
    {
        // Sample JSON array.
        string json = @"[
            { ""Text"": ""First entry"" },
            { ""Text"": ""Second entry"" },
            { ""Text"": ""Third entry"" }
        ]";

        // Deserialize JSON into a list of items.
        List<Item> items = JsonConvert.DeserializeObject<List<Item>>(json) ?? new List<Item>();

        // Create a new blank document.
        Document doc = new Document();

        // Create a block‑level repeating section content control.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(
            doc,
            SdtType.RepeatingSection,
            MarkupLevel.Block)
        {
            Title = "RepeatingSection",
            Tag = "repeating-section"
        };

        // Append the repeating section to the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // For each JSON entry, add a paragraph inside the repeating section.
        foreach (Item item in items)
        {
            Paragraph para = new Paragraph(doc);
            para.AppendChild(new Run(doc, item.Text));
            repeatingSection.AppendChild(para);
        }

        // Save the resulting document.
        doc.Save("RepeatingSectionFromJson.docx");
    }
}
