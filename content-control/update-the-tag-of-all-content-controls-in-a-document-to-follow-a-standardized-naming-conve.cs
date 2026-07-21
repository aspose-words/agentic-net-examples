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
        // Create a sample document with a few content controls.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Sample document with content controls:");

        // Plain text inline content control.
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(sampleDoc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Name",
            Tag = "name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(sampleDoc, "John Doe"));
        builder.InsertNode(plainTextSdt);
        builder.Writeln();

        // Rich text block content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(sampleDoc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Address",
            Tag = "address"
        };
        Paragraph addressParagraph = new Paragraph(sampleDoc);
        addressParagraph.AppendChild(new Run(sampleDoc, "123 Main St"));
        richTextSdt.AppendChild(addressParagraph);
        sampleDoc.FirstSection.Body.AppendChild(richTextSdt);
        builder.Writeln();

        // Checkbox inline content control.
        StructuredDocumentTag checkboxSdt = new StructuredDocumentTag(sampleDoc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Agree",
            Tag = "agree",
            Checked = false
        };
        builder.InsertNode(checkboxSdt);
        builder.Writeln();

        // Save the initial document.
        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath);

        // Load the document for processing.
        Document doc = new Document(inputPath);

        // Prepare a list to hold old and new tag mappings for optional JSON output.
        var tagMappings = new List<object>();

        // Enumerate all StructuredDocumentTag nodes and update their Tag property.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                          .OfType<StructuredDocumentTag>()
                          .ToList();

        for (int i = 0; i < sdtNodes.Count; i++)
        {
            StructuredDocumentTag sdt = sdtNodes[i];
            string oldTag = sdt.Tag ?? string.Empty;
            string newTag = $"Tag_{i + 1}";
            sdt.Tag = newTag;

            tagMappings.Add(new { Index = i + 1, OldTag = oldTag, NewTag = newTag });
        }

        // Save the updated document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Write the tag mapping information to a JSON file.
        string jsonPath = "tag-mapping.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(tagMappings, Formatting.Indented));

        // Optional console output to indicate completion.
        Console.WriteLine($"Processed {sdtNodes.Count} content controls.");
        Console.WriteLine($"Updated document saved as '{outputPath}'.");
        Console.WriteLine($"Tag mapping JSON saved as '{jsonPath}'.");
    }
}
