using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class MergeWordDocumentsWithContentControls
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create first sample document with a plain‑text content control
        Document doc1 = new Document();
        Paragraph para1 = doc1.FirstSection.Body.FirstParagraph;
        StructuredDocumentTag sdt1 = new StructuredDocumentTag(doc1, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "FirstDocumentControl",
            Tag = "first-doc"
        };
        sdt1.RemoveAllChildren();
        sdt1.AppendChild(new Run(doc1, "First document content"));
        para1.AppendChild(sdt1);
        string doc1Path = Path.Combine(outputDir, "doc1.docx");
        doc1.Save(doc1Path);

        // Create second sample document with a plain‑text content control
        Document doc2 = new Document();
        Paragraph para2 = doc2.FirstSection.Body.FirstParagraph;
        StructuredDocumentTag sdt2 = new StructuredDocumentTag(doc2, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SecondDocumentControl",
            Tag = "second-doc"
        };
        sdt2.RemoveAllChildren();
        sdt2.AppendChild(new Run(doc2, "Second document content"));
        para2.AppendChild(sdt2);
        string doc2Path = Path.Combine(outputDir, "doc2.docx");
        doc2.Save(doc2Path);

        // Load the documents for merging
        Document source1 = new Document(doc1Path);
        Document source2 = new Document(doc2Path);

        // Merge source2 into source1
        source1.AppendDocument(source2, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document
        string mergedPath = Path.Combine(outputDir, "merged.docx");
        source1.Save(mergedPath);

        // Export information about each content control to JSON
        var controlInfo = source1.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                 .OfType<StructuredDocumentTag>()
                                 .Select(sdt => new
                                 {
                                     Title = sdt.Title,
                                     Tag = sdt.Tag,
                                     Id = sdt.Id,
                                     Text = sdt.GetText().Trim()
                                 })
                                 .ToList();

        string jsonPath = Path.Combine(outputDir, "contentControls.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(controlInfo, Formatting.Indented));

        // Indicate completion (no interactive input)
        Console.WriteLine($"Merged document saved to: {mergedPath}");
        Console.WriteLine($"Content control details saved to: {jsonPath}");
    }
}
