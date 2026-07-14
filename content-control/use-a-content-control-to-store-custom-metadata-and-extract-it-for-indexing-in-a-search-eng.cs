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
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for the generated files.
        string docPath = Path.Combine(outputDir, "metadata.docx");
        string jsonPath = Path.Combine(outputDir, "metadata.json");

        // -----------------------------------------------------------------
        // 1. Create a document and embed custom metadata in content controls.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom XML part that holds the metadata.
        string xmlContent = "<meta><author>John Doe</author><keywords>example,asp</keywords></meta>";
        CustomXmlPart customXmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xmlContent);

        // Write a heading.
        builder.Writeln("Document Metadata");
        builder.Writeln();

        // Insert an inline plain‑text content control for the author.
        builder.Write("Author: ");
        StructuredDocumentTag authorSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        authorSdt.Title = "Author";
        authorSdt.Tag = "author";
        authorSdt.XmlMapping.SetMapping(customXmlPart, "/meta[1]/author[1]", string.Empty);
        // Add a placeholder run so the SDT is not empty before mapping.
        authorSdt.RemoveAllChildren();
        authorSdt.AppendChild(new Run(doc, "John Doe"));
        builder.CurrentParagraph.AppendChild(authorSdt);
        builder.Writeln();

        // Insert an inline plain‑text content control for the keywords.
        builder.Write("Keywords: ");
        StructuredDocumentTag keywordsSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        keywordsSdt.Title = "Keywords";
        keywordsSdt.Tag = "keywords";
        keywordsSdt.XmlMapping.SetMapping(customXmlPart, "/meta[1]/keywords[1]", string.Empty);
        keywordsSdt.RemoveAllChildren();
        keywordsSdt.AppendChild(new Run(doc, "example,asp"));
        builder.CurrentParagraph.AppendChild(keywordsSdt);
        builder.Writeln();

        // Save the document containing the content controls.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract metadata from the content controls.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        var metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (StructuredDocumentTag sdt in sdtNodes.OfType<StructuredDocumentTag>())
        {
            // Use the Tag property as the key if it is set.
            if (!string.IsNullOrEmpty(sdt.Tag))
            {
                string value = sdt.GetText().Trim();
                metadata[sdt.Tag] = value;
            }
        }

        // Serialize the extracted metadata to JSON.
        string json = JsonConvert.SerializeObject(metadata, Formatting.Indented);
        File.WriteAllText(jsonPath, json);

        // Optional: display the JSON on the console.
        Console.WriteLine("Extracted metadata:");
        Console.WriteLine(json);
    }
}
