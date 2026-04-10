using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several content controls.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Plain text content control.
        StructuredDocumentTag plainText = new StructuredDocumentTag(sampleDoc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PlainTextControl",
            Tag = "PT1"
        };
        builder.InsertNode(plainText);
        builder.Writeln("Plain text content");

        // Rich text content control.
        StructuredDocumentTag richText = new StructuredDocumentTag(sampleDoc, SdtType.RichText, MarkupLevel.Inline)
        {
            Title = "RichTextControl",
            Tag = "RT1"
        };
        builder.InsertNode(richText);
        builder.Writeln("Rich text content");

        // Checkbox content control.
        StructuredDocumentTag checkBox = new StructuredDocumentTag(sampleDoc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "CheckBoxControl",
            Tag = "CB1",
            Checked = true
        };
        builder.InsertNode(checkBox);
        builder.Writeln("Checkbox content");

        // Drop‑down list content control.
        StructuredDocumentTag dropDown = new StructuredDocumentTag(sampleDoc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "DropDownControl",
            Tag = "DD1"
        };
        dropDown.ListItems.Add(new SdtListItem("Option A", "A"));
        dropDown.ListItems.Add(new SdtListItem("Option B", "B"));
        builder.InsertNode(dropDown);
        builder.Writeln("Drop‑down content");

        // Date picker content control.
        StructuredDocumentTag datePicker = new StructuredDocumentTag(sampleDoc, SdtType.Date, MarkupLevel.Inline)
        {
            Title = "DateControl",
            Tag = "DT1"
        };
        builder.InsertNode(datePicker);
        builder.Writeln("Date content");

        // Save the sample document.
        const string samplePath = "SampleWithContentControls.docx";
        sampleDoc.Save(samplePath);

        // Load the document (demonstrates load workflow).
        Document doc = new Document(samplePath);

        // Collect all StructuredDocumentTag nodes.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        List<string> reportLines = new List<string>
        {
            "Content Control Summary Report",
            "==============================",
            $"Total content controls: {sdtNodes.Count}"
        };

        for (int i = 0; i < sdtNodes.Count; i++)
        {
            StructuredDocumentTag sdt = (StructuredDocumentTag)sdtNodes[i];
            string line = $"[{i + 1}] Type: {sdt.SdtType}, Title: \"{sdt.Title}\", Tag: \"{sdt.Tag}\"";
            reportLines.Add(line);
        }

        // Output the report to console.
        foreach (string line in reportLines)
        {
            Console.WriteLine(line);
        }

        // Save the report to a text file.
        const string reportPath = "ContentControlReport.txt";
        File.WriteAllLines(reportPath, reportLines);
    }
}
