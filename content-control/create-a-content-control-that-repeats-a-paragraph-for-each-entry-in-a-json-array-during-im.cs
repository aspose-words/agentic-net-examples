using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Sample JSON array.
        string json = @"[
            { ""Name"": ""Apple"" },
            { ""Name"": ""Banana"" },
            { ""Name"": ""Cherry"" }
        ]";

        // Deserialize JSON into a list of items.
        List<Item> items = JsonConvert.DeserializeObject<List<Item>>(json)!;

        // Create a new blank document.
        Document doc = new Document();

        // Add a title paragraph.
        Paragraph title = new Paragraph(doc);
        title.AppendChild(new Run(doc, "Fruit List"));
        title.ParagraphFormat.StyleName = "Heading 1";
        doc.FirstSection.Body.AppendChild(title);

        // Create a repeating section content control.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(
            doc,
            SdtType.RepeatingSection,
            MarkupLevel.Block)
        {
            Title = "FruitRepeatingSection"
        };

        // Template paragraph inside the repeating section.
        Paragraph templateParagraph = new Paragraph(doc);
        templateParagraph.AppendChild(new Run(doc, "{{Item}}"));
        repeatingSection.AppendChild(templateParagraph);

        // Add the repeating section to the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // For each item, clone the template paragraph, replace the placeholder, and add it as a repeating section item.
        foreach (Item item in items)
        {
            // Create a repeating section item control.
            StructuredDocumentTag itemSdt = new StructuredDocumentTag(
                doc,
                SdtType.RepeatingSectionItem,
                MarkupLevel.Block);

            // Clone the template paragraph.
            Paragraph paraCopy = (Paragraph)templateParagraph.Clone(true);

            // Replace placeholder with actual data.
            paraCopy.Range.Replace("{{Item}}", item.Name);

            // Add the paragraph to the item control.
            itemSdt.AppendChild(paraCopy);

            // Add the item control to the repeating section.
            repeatingSection.AppendChild(itemSdt);
        }

        // Remove the original template paragraph (it has been used for cloning).
        templateParagraph.Remove();

        // Save the resulting document.
        doc.Save("RepeatingSectionFromJson.docx");
    }

    // Simple class matching the JSON structure.
    private class Item
    {
        public string Name { get; set; } = string.Empty;
    }
}
