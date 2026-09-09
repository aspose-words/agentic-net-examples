using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

namespace ContentControlToHtml
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Get the first paragraph of the document (it always exists in a new document).
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

            // ---------- Plain text content control ----------
            StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "customer-name"
            };
            plainTextSdt.RemoveAllChildren();
            plainTextSdt.AppendChild(new Run(doc, "Contoso"));
            paragraph.AppendChild(plainTextSdt);

            // Add a space between controls for readability.
            paragraph.AppendChild(new Run(doc, " "));

            // ---------- Checkbox content control ----------
            StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
            {
                Title = "Agree",
                Tag = "agree",
                Checked = true
            };
            paragraph.AppendChild(checkBoxSdt);

            // Add a space between controls.
            paragraph.AppendChild(new Run(doc, " "));

            // ---------- Drop‑down list content control ----------
            StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
            {
                Title = "Options",
                Tag = "options"
            };
            dropDownSdt.ListItems.Add(new SdtListItem("Option A", "A"));
            dropDownSdt.ListItems.Add(new SdtListItem("Option B", "B"));
            paragraph.AppendChild(dropDownSdt);

            // Save the document as HTML. The default behavior exports content controls as data‑attributes.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "ContentControls.html");
            doc.Save(htmlPath, htmlOptions);

            // Indicate completion.
            Console.WriteLine($"Document converted to HTML with data‑attributes saved at: {htmlPath}");
        }
    }
}
