using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the RTF template.
        var loadOptions = new RtfLoadOptions();               // default load options
        Document doc = new Document("Template.rtf", loadOptions);

        // Create a DocumentBuilder for editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a placeholder in the template that will be replaced by the HTML.
        const string placeholder = "<<HTML>>";

        // Remove the placeholder text and position the builder at the end of the document.
        // (Replace returns the number of replacements made; we only need to know it succeeded.)
        if (doc.Range.Replace(placeholder, string.Empty) > 0)
        {
            builder.MoveToDocumentEnd();
        }

        // HTML content to insert dynamically.
        const string html = "<h1 style='color:blue;'>Dynamic Title</h1>" +
                            "<p>This paragraph is inserted from HTML.</p>";

        // Insert the HTML. Use builder formatting as the base formatting.
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting);

        // Save the modified document back to RTF.
        var saveOptions = new RtfSaveOptions();               // default save options
        doc.Save("Result.rtf", saveOptions);
    }
}
