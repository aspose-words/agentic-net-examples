using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create first sample document with an inline plain‑text content control
        Document doc1 = new Document();
        Paragraph para1 = doc1.FirstSection.Body.FirstParagraph;
        StructuredDocumentTag sdt1 = new StructuredDocumentTag(doc1, SdtType.PlainText, MarkupLevel.Inline);
        sdt1.Title = "FirstName";
        sdt1.Tag = "first-name";
        sdt1.RemoveAllChildren();
        sdt1.AppendChild(new Run(doc1, "John"));
        para1.AppendChild(sdt1);
        string doc1Path = Path.Combine(outputDir, "doc1.docx");
        doc1.Save(doc1Path);

        // Create second sample document with a block‑level rich‑text content control
        Document doc2 = new Document();
        StructuredDocumentTag sdt2 = new StructuredDocumentTag(doc2, SdtType.RichText, MarkupLevel.Block);
        sdt2.Title = "Address";
        sdt2.Tag = "address";
        Paragraph innerPara = new Paragraph(doc2);
        innerPara.AppendChild(new Run(doc2, "123 Main St, Anytown"));
        sdt2.AppendChild(innerPara);
        doc2.FirstSection.Body.AppendChild(sdt2);
        string doc2Path = Path.Combine(outputDir, "doc2.docx");
        doc2.Save(doc2Path);

        // Load the documents for merging
        Document mainDoc = new Document(doc1Path);
        Document secondDoc = new Document(doc2Path);

        // Append the second document to the first one
        mainDoc.AppendDocument(secondDoc, ImportFormatMode.KeepSourceFormatting);

        // Update IDs of all content controls to ensure uniqueness after merge
        // The Id property is read‑only, so we set the underlying private field via reflection.
        var allSdt = mainDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                            .OfType<StructuredDocumentTag>()
                            .ToList();

        // Attempt to locate the private field that stores the Id (commonly named "_id" or "mId").
        // If the field name changes in future versions, this reflection may need adjustment.
        FieldInfo idField = typeof(StructuredDocumentTag)
            .GetFields(BindingFlags.Instance | BindingFlags.NonPublic)
            .FirstOrDefault(f => f.Name.Equals("_id", StringComparison.OrdinalIgnoreCase) ||
                                 f.Name.Equals("mId", StringComparison.OrdinalIgnoreCase) ||
                                 f.Name.Equals("id", StringComparison.OrdinalIgnoreCase));

        if (idField != null)
        {
            foreach (var sdt in allSdt)
            {
                idField.SetValue(sdt, Guid.NewGuid().ToString());
            }
        }

        // Save the merged document
        string mergedPath = Path.Combine(outputDir, "merged.docx");
        mainDoc.Save(mergedPath);

        // Export information about the content controls to JSON for verification
        var sdtInfo = allSdt.Select(s => new
        {
            Title = s.Title,
            Tag = s.Tag,
            Id = s.Id,
            Text = s.GetText().Trim()
        }).ToList();

        string jsonPath = Path.Combine(outputDir, "merged_content_controls.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sdtInfo, Formatting.Indented));
    }
}
