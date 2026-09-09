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
            { ""Name"": ""Alice"" },
            { ""Name"": ""Bob"" },
            { ""Name"": ""Charlie"" }
        ]";

        // Deserialize JSON into a list of simple objects.
        List<Dictionary<string, string>> items = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(json);

        // Create a new blank document.
        Document doc = new Document();

        // Create a block‑level repeating section content control.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Template paragraph that will be cloned for each JSON entry.
        Paragraph templateParagraph = new Paragraph(doc);
        templateParagraph.AppendChild(new Run(doc, "Placeholder"));

        // For each item in the JSON array, create a repeating section item and insert a populated paragraph.
        foreach (Dictionary<string, string> entry in items)
        {
            // Clone the template paragraph.
            Paragraph paraClone = (Paragraph)templateParagraph.Clone(true);

            // Replace placeholder text with the actual value from JSON.
            if (paraClone.Runs.Count > 0 && entry.TryGetValue("Name", out string name))
            {
                paraClone.Runs[0].Text = name;
            }

            // Create a repeating section item and add the populated paragraph to it.
            StructuredDocumentTag itemSdt = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
            itemSdt.AppendChild(paraClone);

            // Append the item to the repeating section.
            repeatingSection.AppendChild(itemSdt);
        }

        // Save the resulting document.
        doc.Save("RepeatingSectionFromJson.docx");
    }
}
