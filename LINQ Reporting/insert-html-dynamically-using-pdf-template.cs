using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertHtmlIntoPdfTemplate
{
    static void Main()
    {
        // Path to the PDF template file.
        string templatePath = @"C:\Templates\Template.pdf";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Output\Result.pdf";

        // Load the PDF template into an Aspose.Words Document.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: move the cursor to a bookmark named "Content" where HTML will be inserted.
        // If the bookmark does not exist, the builder will stay at the current position (document start).
        if (doc.Range.Bookmarks["Content"] != null)
            builder.MoveToBookmark("Content");

        // HTML string to be inserted dynamically.
        string html = @"
            <h2 style='color:#2E86C1;'>Dynamic Section</h2>
            <p>This paragraph was inserted from an HTML fragment at " + DateTime.Now.ToString("f") + @"</p>
            <ul>
                <li>Item 1</li>
                <li>Item 2</li>
                <li>Item 3</li>
            </ul>";

        // Insert the HTML using builder formatting and remove the trailing empty paragraph.
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

        // Configure PDF save options if needed (e.g., preserve form fields, set compliance level, etc.).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: ensure that any form fields in the template remain interactive.
            PreserveFormFields = true,
            // Example: set PDF/A-1b compliance.
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the modified document as a PDF using the specified options.
        doc.Save(outputPath, pdfOptions);
    }
}
