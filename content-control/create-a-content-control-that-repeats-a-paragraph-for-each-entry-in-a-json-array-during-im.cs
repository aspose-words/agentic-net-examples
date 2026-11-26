using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Sample JSON array to import.
        string json = "[\"First item\",\"Second item\",\"Third item\"]";

        // Deserialize JSON array into a list of strings.
        List<string> items = JsonConvert.DeserializeObject<List<string>>(json) ?? new List<string>();

        // Create a new blank document.
        Document doc = new Document();

        // Create a block‑level repeating section content control.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "ItemsRepeatingSection",
            Tag = "items-repeating"
        };

        // For each entry in the JSON array, add a paragraph inside the repeating section.
        foreach (string item in items)
        {
            Paragraph para = new Paragraph(doc);
            para.AppendChild(new Run(doc, item));
            repeatingSection.AppendChild(para);
        }

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the resulting document.
        string outputPath = "RepeatingSection.docx";
        doc.Save(outputPath);
    }
}
