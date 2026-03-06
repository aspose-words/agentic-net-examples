using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertHtmlDynamic
{
    static void Main()
    {
        // Load a DOTX template. The template can contain placeholders or predefined styles.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Switch that determines whether to apply the builder's formatting to the inserted HTML.
        // When true, the formatting from the builder (styles, fonts, etc.) is used as base formatting.
        // When false, the HTML is inserted with its own default formatting.
        bool sourceStyles = true; // Change this value to toggle the behavior.

        // Build the HTML string dynamically. This example changes the text color based on the switch.
        string html = $"<p style='color:{(sourceStyles ? "red" : "blue")}'>Dynamic HTML content inserted at {DateTime.Now}</p>";

        // Choose the appropriate HtmlInsertOptions based on the sourceStyles switch.
        HtmlInsertOptions insertOptions = sourceStyles
            ? HtmlInsertOptions.UseBuilderFormatting
            : HtmlInsertOptions.None;

        // Insert the HTML into the document using the selected options.
        builder.InsertHtml(html, insertOptions);

        // Save the resulting document.
        doc.Save("Result.docx", SaveOptions.CreateSaveOptions(SaveFormat.Docx));
    }
}
