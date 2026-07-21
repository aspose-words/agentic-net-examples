using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists (current working directory)
        string outputDir = Directory.GetCurrentDirectory();

        // -------------------------------------------------
        // Create first sample document with a plain‑text content control
        // -------------------------------------------------
        Document doc1 = new Document();
        Paragraph para1 = doc1.FirstSection.Body.FirstParagraph;

        StructuredDocumentTag sdtPlain = new StructuredDocumentTag(doc1, SdtType.PlainText, MarkupLevel.Inline);
        sdtPlain.Title = "FirstName";
        sdtPlain.Tag = "first-name";
        sdtPlain.RemoveAllChildren();
        sdtPlain.AppendChild(new Run(doc1, "John"));
        para1.AppendChild(sdtPlain);

        doc1.Save(Path.Combine(outputDir, "doc1.docx"));

        // -------------------------------------------------
        // Create second sample document with a rich‑text content control
        // -------------------------------------------------
        Document doc2 = new Document();
        Paragraph para2 = doc2.FirstSection.Body.FirstParagraph;

        StructuredDocumentTag sdtRich = new StructuredDocumentTag(doc2, SdtType.RichText, MarkupLevel.Block);
        sdtRich.Title = "Address";
        sdtRich.Tag = "address";

        Paragraph innerPara = new Paragraph(doc2);
        innerPara.AppendChild(new Run(doc2, "123 Main St, Anytown"));
        sdtRich.AppendChild(innerPara);

        doc2.FirstSection.Body.AppendChild(sdtRich);
        doc2.Save(Path.Combine(outputDir, "doc2.docx"));

        // -------------------------------------------------
        // Load the documents for merging
        // -------------------------------------------------
        Document mainDoc = new Document(Path.Combine(outputDir, "doc1.docx"));
        Document otherDoc = new Document(Path.Combine(outputDir, "doc2.docx"));

        // Append the second document to the first one
        mainDoc.AppendDocument(otherDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // Gather all content controls (StructuredDocumentTag nodes)
        // -------------------------------------------------
        var allSdt = mainDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                            .OfType<StructuredDocumentTag>()
                            .ToList();

        // -------------------------------------------------
        // Serialize content‑control metadata to JSON
        // -------------------------------------------------
        var metadata = allSdt.Select(s => new
        {
            Title = s.Title,
            Tag = s.Tag,
            Id = s.Id,               // Id is read‑only and already unique
            Text = s.GetText().Trim()
        }).ToList();

        string json = JsonConvert.SerializeObject(metadata, Formatting.Indented);
        File.WriteAllText(Path.Combine(outputDir, "contentControls.json"), json);

        // -------------------------------------------------
        // Save the merged document
        // -------------------------------------------------
        mainDoc.Save(Path.Combine(outputDir, "Merged.docx"));
    }
}
