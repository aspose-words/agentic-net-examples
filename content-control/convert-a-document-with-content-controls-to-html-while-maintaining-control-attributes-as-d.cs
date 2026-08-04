using System;
using System.IO;
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

        // Add an introductory paragraph.
        builder.Writeln("Document with various content controls:");

        // -----------------------------------------------------------------
        // Inline plain‑text content control.
        // -----------------------------------------------------------------
        Paragraph inlineParagraph = doc.FirstSection.Body.FirstParagraph;

        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "John Doe"));
        inlineParagraph.AppendChild(new Run(doc, " Name: "));
        inlineParagraph.AppendChild(plainTextSdt);

        // -----------------------------------------------------------------
        // Inline checkbox content control.
        // -----------------------------------------------------------------
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Agree",
            Tag = "agree",
            Checked = true
        };
        inlineParagraph.AppendChild(new Run(doc, " "));
        inlineParagraph.AppendChild(checkBoxSdt);
        inlineParagraph.AppendChild(new Run(doc, " I agree"));

        // -----------------------------------------------------------------
        // Inline drop‑down list content control.
        // -----------------------------------------------------------------
        StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Country",
            Tag = "country"
        };
        dropDownSdt.ListItems.Add(new SdtListItem("USA", "USA"));
        dropDownSdt.ListItems.Add(new SdtListItem("Canada", "Canada"));
        dropDownSdt.ListItems.Add(new SdtListItem("Mexico", "Mexico"));
        dropDownSdt.RemoveAllChildren();
        dropDownSdt.AppendChild(new Run(doc, "USA"));
        inlineParagraph.AppendChild(new Run(doc, " "));
        inlineParagraph.AppendChild(dropDownSdt);
        inlineParagraph.AppendChild(new Run(doc, " (Select country)"));

        // -----------------------------------------------------------------
        // Block‑level rich‑text content control.
        // -----------------------------------------------------------------
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Comments",
            Tag = "comments"
        };
        Paragraph innerParagraph = new Paragraph(doc);
        innerParagraph.AppendChild(new Run(doc, "This is a sample comment inside a rich‑text content control."));
        richTextSdt.AppendChild(innerParagraph);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the source DOCX (optional, demonstrates that the file exists).
        const string docxPath = "ContentControls.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Convert the document to HTML while preserving content‑control attributes.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // The default behavior exports content controls as <span> elements
            // with data‑attributes that retain Title, Tag, Id, etc.
            // No additional configuration is required for this scenario.
        };

        const string htmlPath = "ContentControls.html";
        doc.Save(htmlPath, htmlOptions);

        // Inform the user (console output is allowed as it does not require input).
        Console.WriteLine($"DOCX saved to: {Path.GetFullPath(docxPath)}");
        Console.WriteLine($"HTML saved to: {Path.GetFullPath(htmlPath)}");
    }
}
