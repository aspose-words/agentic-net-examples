using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control.
        StructuredDocumentTag plainSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainSdt.RemoveAllChildren();
        plainSdt.AppendChild(new Run(doc, "John Doe"));
        builder.InsertNode(plainSdt);
        builder.Writeln(); // separate from next control

        // Insert a checkbox content control.
        StructuredDocumentTag checkSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Agree",
            Tag = "agree",
            Checked = false
        };
        builder.InsertNode(checkSdt);
        builder.Writeln();

        // Insert a drop‑down list content control.
        StructuredDocumentTag dropSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Country",
            Tag = "country"
        };
        dropSdt.ListItems.Add(new SdtListItem("USA", "US"));
        dropSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
        builder.InsertNode(dropSdt);
        builder.Writeln();

        // Save the source DOCX (optional, just for reference).
        const string docxPath = "sample.docx";
        doc.Save(docxPath);

        // Convert the document to HTML. Content‑control attributes (Title, Tag) are exported as data‑attributes by default.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        const string htmlPath = "sample.html";
        doc.Save(htmlPath, htmlOptions);

        // Output a short confirmation.
        Console.WriteLine($"Document saved as '{docxPath}'.");
        Console.WriteLine($"HTML conversion saved as '{htmlPath}'.");
    }
}
