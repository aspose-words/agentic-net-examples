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
        // Create first sample document with a plain‑text content control.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Document 1 – start");
        StructuredDocumentTag sdt1 = new StructuredDocumentTag(doc1, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PersonName",
            Tag = "person-name"
        };
        sdt1.RemoveAllChildren();
        sdt1.AppendChild(new Run(doc1, "Alice"));
        builder1.CurrentParagraph.AppendChild(sdt1);
        builder1.Writeln();
        builder1.Writeln("Document 1 – end");
        doc1.Save("doc1.docx");

        // Create second sample document with a rich‑text content control.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Document 2 – start");
        StructuredDocumentTag sdt2 = new StructuredDocumentTag(doc2, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "AddressBlock",
            Tag = "address-block"
        };
        Paragraph para = new Paragraph(doc2);
        para.AppendChild(new Run(doc2, "123 Main St.\nCityville"));
        sdt2.AppendChild(para);
        doc2.FirstSection.Body.AppendChild(sdt2);
        builder2.Writeln();
        builder2.Writeln("Document 2 – end");
        doc2.Save("doc2.docx");

        // Load the documents for merging.
        Document mainDoc = new Document("doc1.docx");
        Document secondaryDoc = new Document("doc2.docx");

        // Merge secondaryDoc into mainDoc.
        mainDoc.AppendDocument(secondaryDoc, ImportFormatMode.KeepSourceFormatting);

        // Gather all content controls.
        List<StructuredDocumentTag> allSdt = mainDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .ToList();

        // Save the merged document.
        string mergedPath = "merged.docx";
        mainDoc.Save(mergedPath);

        // Export information about the content controls to JSON.
        var payload = allSdt
            .Select((sdt, index) => new
            {
                Title = sdt.Title ?? string.Empty,
                Tag = sdt.Tag ?? string.Empty,
                // Use a sequential identifier for the report (original Id is read‑only).
                Id = index + 1,
                Text = sdt.GetText().Trim()
            })
            .ToList();

        string json = JsonConvert.SerializeObject(payload, Formatting.Indented);
        File.WriteAllText("merged_sdt.json", json);
    }
}
