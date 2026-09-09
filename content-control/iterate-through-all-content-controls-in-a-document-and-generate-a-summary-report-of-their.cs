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
        // Create a sample document with various content controls.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Ensure the document has at least one paragraph.
        builder.Writeln("Document with several content controls:");
        Paragraph firstParagraph = sampleDoc.FirstSection.Body.FirstParagraph;

        // Inline plain‑text content control.
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(sampleDoc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(sampleDoc, "Contoso"));
        firstParagraph.AppendChild(plainTextSdt);
        firstParagraph.AppendChild(new Run(sampleDoc, " "));

        // Inline checkbox content control.
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(sampleDoc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AcceptTerms",
            Tag = "accept-terms",
            Checked = true
        };
        firstParagraph.AppendChild(checkBoxSdt);
        firstParagraph.AppendChild(new Run(sampleDoc, " "));

        // Inline drop‑down list content control.
        StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(sampleDoc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Country",
            Tag = "country"
        };
        dropDownSdt.ListItems.Add(new SdtListItem("USA", "US"));
        dropDownSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
        firstParagraph.AppendChild(dropDownSdt);
        firstParagraph.AppendChild(new Run(sampleDoc, " "));

        // Inline date picker content control.
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(sampleDoc, SdtType.Date, MarkupLevel.Inline)
        {
            Title = "BirthDate",
            Tag = "birth-date",
            DateDisplayFormat = "yyyy-MM-dd"
        };
        firstParagraph.AppendChild(dateSdt);
        firstParagraph.AppendChild(new Run(sampleDoc, " "));

        // Block‑level rich‑text content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(sampleDoc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Comments",
            Tag = "comments"
        };
        Paragraph blockParagraph = new Paragraph(sampleDoc);
        blockParagraph.AppendChild(new Run(sampleDoc, "Enter your comments here."));
        richTextSdt.AppendChild(blockParagraph);
        sampleDoc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the sample document.
        const string samplePath = "sample.docx";
        sampleDoc.Save(samplePath);

        // Load the document (simulating a separate processing step).
        Document doc = new Document(samplePath);

        // Collect information about each content control.
        var controlsInfo = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                              .OfType<StructuredDocumentTag>()
                              .Select(sdt => new
                              {
                                  Type = sdt.SdtType.ToString(),
                                  Title = sdt.Title,
                                  Tag = sdt.Tag
                              })
                              .ToList();

        // Serialize the summary to JSON.
        string jsonReport = JsonConvert.SerializeObject(controlsInfo, Formatting.Indented);
        const string reportPath = "content_controls_report.json";
        File.WriteAllText(reportPath, jsonReport);

        // Output the JSON to the console (no interactive prompts).
        Console.WriteLine("Content Control Summary:");
        Console.WriteLine(jsonReport);
    }
}
