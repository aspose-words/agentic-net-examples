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

        // Get the first paragraph of the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // ----- Plain text content control -----
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        plainTextSdt.Title = "CustomerName";
        plainTextSdt.Tag = "customer-name";
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "Contoso"));
        paragraph.AppendChild(plainTextSdt);

        // Add a space between controls.
        paragraph.AppendChild(new Run(doc, " "));

        // ----- Checkbox content control -----
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
        checkBoxSdt.Title = "Agree";
        checkBoxSdt.Tag = "agree";
        checkBoxSdt.Checked = true;
        paragraph.AppendChild(checkBoxSdt);

        // Save the source DOCX file.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document that contains the content controls.
        Document loadedDoc = new Document(inputPath);

        // Configure HTML save options.
        // The ExportContentControlsAsDataAttributes property is not available in this version of Aspose.Words,
        // so we simply use the default behavior which includes content control information in the HTML.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save the document as HTML.
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "output.html");
        loadedDoc.Save(htmlPath, htmlOptions);

        // Indicate that the conversion has finished.
        Console.WriteLine("Conversion completed. HTML saved to: " + htmlPath);
    }
}
