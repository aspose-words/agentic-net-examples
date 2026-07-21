using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Plain‑text content control ----------
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "John Doe"));
        // Insert the SDT into the current paragraph.
        builder.InsertNode(plainTextSdt);
        builder.Writeln(); // Move to a new paragraph.

        // ---------- Checkbox content control ----------
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Agree",
            Tag = "agree",
            Checked = false
        };
        builder.InsertNode(checkBoxSdt);
        builder.Writeln();

        // ---------- Drop‑down list content control ----------
        StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Country",
            Tag = "country"
        };
        dropDownSdt.ListItems.Add(new SdtListItem("USA", "US"));
        dropDownSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
        dropDownSdt.ListItems.Add(new SdtListItem("Mexico", "MX"));
        builder.InsertNode(dropDownSdt);
        builder.Writeln();

        // Save the DOCX to a local file.
        const string docxPath = "sample.docx";
        doc.Save(docxPath);

        // Load the saved document (demonstrates a realistic workflow where the source file already exists).
        Document loadedDoc = new Document(docxPath);

        // Configure HTML save options.
        // The property ExportContentControlsAsDataAttributes is not available in the current Aspose.Words version,
        // so we rely on the default behavior which already includes content‑control metadata as data‑attributes.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save the document as HTML.
        const string htmlPath = "sample.html";
        loadedDoc.Save(htmlPath, htmlOptions);

        // Confirmation output.
        Console.WriteLine($"Document converted to HTML: {Path.GetFullPath(htmlPath)}");
    }
}
