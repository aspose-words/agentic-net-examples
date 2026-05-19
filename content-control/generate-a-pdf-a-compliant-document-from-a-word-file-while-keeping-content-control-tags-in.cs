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
        // Create a sample DOCX with several content controls.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Plain‑text content control.
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PlainTextControl",
            Tag = "plain-text"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "Enter name"));
        builder.InsertNode(plainTextSdt);
        builder.Writeln();

        // Checkbox content control.
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AgreeCheck",
            Tag = "agree-check",
            Checked = false
        };
        builder.InsertNode(checkBoxSdt);
        builder.Writeln();

        // Drop‑down list content control.
        StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "CountrySelect",
            Tag = "country-select"
        };
        dropDownSdt.ListItems.Add(new SdtListItem("USA", "US"));
        dropDownSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
        dropDownSdt.ListItems.Add(new SdtListItem("Mexico", "MX"));
        builder.InsertNode(dropDownSdt);
        builder.Writeln();

        // Save the DOCX to the working directory.
        const string docxPath = "sample.docx";
        doc.Save(docxPath);

        // Load the document back.
        Document loadedDoc = new Document(docxPath);

        // Example modification: set the plain‑text control's content.
        StructuredDocumentTag? foundPlainText = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(sdt => sdt.Title == "PlainTextControl");
        if (foundPlainText != null)
        {
            foundPlainText.RemoveAllChildren();
            foundPlainText.AppendChild(new Run(loadedDoc, "John Doe"));
        }

        // Export metadata of all content controls to JSON (demonstrates Newtonsoft.Json usage).
        var sdtInfo = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Select(sdt => new
            {
                sdt.Title,
                sdt.Tag,
                Type = sdt.SdtType.ToString(),
                Text = sdt.GetText().Trim()
            })
            .ToList();

        string json = JsonConvert.SerializeObject(sdtInfo, Formatting.Indented);
        File.WriteAllText("content-controls.json", json);

        // Configure PDF/A‑1a compliance and preserve content controls as form fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a,
            PreserveFormFields = true,
            UseSdtTagAsFormFieldName = true,
            // ExportDocumentStructure is required for PDF/A‑1a, but the property is ignored automatically.
            ExportDocumentStructure = true
        };

        // Save the document as PDF/A.
        const string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, pdfOptions);
    }
}
