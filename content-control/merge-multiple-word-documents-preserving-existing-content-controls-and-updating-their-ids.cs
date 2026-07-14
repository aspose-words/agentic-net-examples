using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string outputDir = Directory.GetCurrentDirectory();

        // -------------------------------------------------------------
        // 1. Create the first source document with a plain‑text SDT.
        // -------------------------------------------------------------
        Document doc1 = new Document();
        StructuredDocumentTag sdt1 = new StructuredDocumentTag(doc1, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "FirstControl",
            Tag = "FirstTag"
        };
        sdt1.RemoveAllChildren();
        sdt1.AppendChild(new Run(doc1, "Content from document 1"));
        doc1.FirstSection.Body.FirstParagraph.AppendChild(sdt1);
        string doc1Path = Path.Combine(outputDir, "doc1.docx");
        doc1.Save(doc1Path, SaveFormat.Docx);

        // -------------------------------------------------------------
        // 2. Create the second source document with a rich‑text SDT.
        // -------------------------------------------------------------
        Document doc2 = new Document();
        StructuredDocumentTag sdt2 = new StructuredDocumentTag(doc2, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "SecondControl",
            Tag = "SecondTag"
        };
        Paragraph para = new Paragraph(doc2);
        para.AppendChild(new Run(doc2, "Content from document 2"));
        sdt2.AppendChild(para);
        doc2.FirstSection.Body.AppendChild(sdt2);
        string doc2Path = Path.Combine(outputDir, "doc2.docx");
        doc2.Save(doc2Path, SaveFormat.Docx);

        // -------------------------------------------------------------
        // 3. Load the source documents and merge them.
        // -------------------------------------------------------------
        Document destination = new Document(doc1Path);
        Document sourceToAppend = new Document(doc2Path);
        destination.AppendDocument(sourceToAppend, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------------------
        // 4. Assign a new integer CustomNodeId to each content control.
        //    This ensures uniqueness after the merge.
        // -------------------------------------------------------------
        var allControls = destination
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .ToList();

        var random = new Random();
        foreach (var control in allControls)
        {
            // CustomNodeId is an integer in this version of Aspose.Words.
            control.CustomNodeId = random.Next(int.MinValue, int.MaxValue);
        }

        // -------------------------------------------------------------
        // 5. Save the merged document.
        // -------------------------------------------------------------
        string mergedPath = Path.Combine(outputDir, "merged.docx");
        destination.Save(mergedPath, SaveFormat.Docx);

        // -------------------------------------------------------------
        // 6. (Optional) Export metadata of the content controls to JSON.
        // -------------------------------------------------------------
        var metadata = allControls.Select(c => new
        {
            PersistentId = c.Id,
            CustomNodeId = c.CustomNodeId,
            Title = c.Title,
            Tag = c.Tag,
            Text = c.GetText().Trim()
        }).ToList();

        string json = JsonConvert.SerializeObject(metadata, Formatting.Indented);
        File.WriteAllText(Path.Combine(outputDir, "merged.json"), json);
    }
}
