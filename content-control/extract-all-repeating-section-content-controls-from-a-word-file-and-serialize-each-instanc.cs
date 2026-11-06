using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with repeating section content controls.
        Document doc = new Document();

        // First repeating section.
        StructuredDocumentTag repeating1 = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeating1.Title = "Products";
        repeating1.Tag = "products-section";

        Paragraph p1 = new Paragraph(doc);
        p1.AppendChild(new Run(doc, "Product A"));
        repeating1.AppendChild(p1);

        // Add a second item inside the same repeating section to demonstrate multiple items.
        Paragraph p2 = new Paragraph(doc);
        p2.AppendChild(new Run(doc, "Product B"));
        repeating1.AppendChild(p2);

        // Insert the first repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeating1);

        // Second repeating section with different content.
        StructuredDocumentTag repeating2 = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeating2.Title = "Services";
        repeating2.Tag = "services-section";

        Paragraph p3 = new Paragraph(doc);
        p3.AppendChild(new Run(doc, "Consulting"));
        repeating2.AppendChild(p3);

        doc.FirstSection.Body.AppendChild(repeating2);

        // Save the sample document.
        const string inputPath = "sample.docx";
        doc.Save(inputPath);

        // Load the document from file.
        Document loadedDoc = new Document(inputPath);

        // Extract all repeating section content controls.
        List<object> repeatingSections = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection)
            .Select(sdt => new
            {
                Title = sdt.Title,
                Tag = sdt.Tag,
                Text = sdt.GetText().Trim()
            })
            .Cast<object>()
            .ToList();

        // Serialize the extracted data to JSON.
        string json = JsonConvert.SerializeObject(repeatingSections, Formatting.Indented);

        // Write JSON to file.
        const string outputPath = "repeating-sections.json";
        File.WriteAllText(outputPath, json);
    }
}
