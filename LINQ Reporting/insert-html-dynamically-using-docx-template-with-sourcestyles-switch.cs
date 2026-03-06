using System;
using Aspose.Words;

class InsertHtmlDynamic
{
    static void Main()
    {
        // Path to the DOCX template that contains placeholders or bookmarks.
        string templatePath = @"C:\Docs\Template.docx";

        // Path where the final document will be saved.
        string outputPath = @"C:\Docs\Result.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder to work with the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example HTML fragment that will be inserted dynamically.
        string html = "<h1 style='color:blue;'>Dynamic Title</h1>" +
                      "<p>This paragraph is <b>bold</b> and <i>italic</i>.</p>";

        // Switch that decides whether to keep the HTML's own styling.
        // true  -> keep source HTML styles (useBuilderFormatting = false)
        // false -> apply the builder's current formatting to the inserted HTML (useBuilderFormatting = true)
        bool keepSourceStyles = true;

        if (keepSourceStyles)
        {
            // Insert HTML while preserving its own formatting.
            // The second parameter (useBuilderFormatting) is set to false.
            builder.InsertHtml(html, false);
        }
        else
        {
            // Insert HTML and let the builder's formatting be the base.
            // The second parameter (useBuilderFormatting) is set to true.
            builder.InsertHtml(html, true);
        }

        // Save the resulting document.
        doc.Save(outputPath);
    }
}
