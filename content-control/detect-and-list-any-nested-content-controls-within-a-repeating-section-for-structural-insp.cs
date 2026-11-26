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
        // Create a new document.
        Document doc = new Document();

        // Create a repeating section content control (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "RepeatingSection",
            Tag = "repeating-section"
        };

        // First nested plain text content control (inline).
        Paragraph paraPlain = new Paragraph(doc);
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NestedPlain",
            Tag = "nested-plain"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "Nested plain text"));
        paraPlain.AppendChild(plainTextSdt);
        repeatingSection.AppendChild(paraPlain);

        // Second nested checkbox content control (inline).
        Paragraph paraCheckbox = new Paragraph(doc);
        StructuredDocumentTag checkboxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "NestedCheckbox",
            Tag = "nested-checkbox",
            Checked = true
        };
        paraCheckbox.AppendChild(checkboxSdt);
        repeatingSection.AppendChild(paraCheckbox);

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the sample document.
        doc.Save("sample.docx");

        // Detect nested content controls within repeating sections.
        List<object> nestedInfo = new List<object>();

        IEnumerable<StructuredDocumentTag> repeatingControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection);

        foreach (StructuredDocumentTag repeating in repeatingControls)
        {
            // Find all descendant content controls inside this repeating section.
            IEnumerable<StructuredDocumentTag> nestedControls = repeating.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>();

            foreach (StructuredDocumentTag nested in nestedControls)
            {
                var info = new
                {
                    ParentTitle = repeating.Title,
                    ParentTag = repeating.Tag,
                    NestedTitle = nested.Title,
                    NestedTag = nested.Tag,
                    NestedType = nested.SdtType.ToString()
                };
                nestedInfo.Add(info);
                Console.WriteLine($"Parent: {repeating.Title} ({repeating.Tag}) -> Nested: {nested.Title} ({nested.Tag}) Type: {nested.SdtType}");
            }
        }

        // Serialize the inspection result to JSON.
        string json = JsonConvert.SerializeObject(nestedInfo, Formatting.Indented);
        File.WriteAllText("nested-sdts.json", json);
    }
}
