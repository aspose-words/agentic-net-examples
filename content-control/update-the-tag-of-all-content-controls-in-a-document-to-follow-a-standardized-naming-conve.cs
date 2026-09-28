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
        // Create a sample document with several content controls.
        Document seedDoc = new Document();
        Paragraph firstPara = seedDoc.FirstSection.Body.FirstParagraph;

        // Plain text content control.
        StructuredDocumentTag plainSdt = new StructuredDocumentTag(seedDoc, SdtType.PlainText, MarkupLevel.Inline);
        plainSdt.Title = "Customer Name";
        plainSdt.Tag = "custName";
        plainSdt.RemoveAllChildren();
        plainSdt.AppendChild(new Run(seedDoc, "Alice"));
        firstPara.AppendChild(plainSdt);

        // Rich text content control.
        StructuredDocumentTag richSdt = new StructuredDocumentTag(seedDoc, SdtType.RichText, MarkupLevel.Block);
        richSdt.Title = "Address Block";
        richSdt.Tag = "addrBlk";
        Paragraph richPara = new Paragraph(seedDoc);
        richPara.AppendChild(new Run(seedDoc, "123 Main St"));
        richSdt.AppendChild(richPara);
        seedDoc.FirstSection.Body.AppendChild(richSdt);

        // Drop-down list content control.
        StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(seedDoc, SdtType.DropDownList, MarkupLevel.Inline);
        dropDownSdt.Title = "Country Selector";
        dropDownSdt.Tag = "cntSel";
        dropDownSdt.ListItems.Add(new SdtListItem("USA", "US"));
        dropDownSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
        firstPara.AppendChild(dropDownSdt);

        // Save the seed document.
        const string inputPath = "input.docx";
        seedDoc.Save(inputPath);

        // Load the document for processing.
        Document doc = new Document(inputPath);

        // Enumerate all content controls and update their Tag property.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                          .OfType<StructuredDocumentTag>();

        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Standardized tag: Title with spaces replaced by underscores, suffixed with "_Tag".
            string baseTitle = sdt.Title ?? "Unnamed";
            string standardizedTag = $"{baseTitle.Replace(' ', '_')}_Tag";
            sdt.Tag = standardizedTag;
        }

        // Save the updated document.
        const string outputPath = "updated.docx";
        doc.Save(outputPath);

        // Optional: write a JSON report of the updated tags.
        var report = sdtNodes.Select(s => new
        {
            Title = s.Title,
            Tag = s.Tag,
            Type = s.SdtType.ToString()
        }).ToList();

        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("tags-report.json", json);
    }
}
