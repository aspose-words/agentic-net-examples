using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDemo
{
    public class HtmlInserter
    {
        /// <summary>
        /// Inserts an HTML fragment into a DOCM template.
        /// </summary>
        /// <param name="templatePath">Path to the DOCM template.</param>
        /// <param name="outputPath">Path where the resulting document will be saved.</param>
        /// <param name="html">HTML string to insert.</param>
        /// <param name="useSourceStyles">
        /// If true, the HTML is inserted using the formatting defined in the HTML itself (source styles).
        /// If false, the formatting of the DocumentBuilder (destination styles) is applied.
        /// </param>
        public static void InsertHtmlIntoTemplate(string templatePath, string outputPath, string html, bool useSourceStyles)
        {
            // Load the DOCM template.
            Document doc = new Document(templatePath);

            // Create a DocumentBuilder for the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to a bookmark named "InsertHere" if it exists;
            // otherwise, move to the end of the document.
            if (doc.Range.Bookmarks["InsertHere"] != null)
                builder.MoveToBookmark("InsertHere");
            else
                builder.MoveToDocumentEnd();

            // Insert the HTML using the appropriate overload based on the switch.
            if (useSourceStyles)
            {
                // Keep the HTML's own formatting.
                builder.InsertHtml(html);
            }
            else
            {
                // Apply the builder's current formatting as base formatting.
                builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting);
            }

            // Save the resulting document.
            doc.Save(outputPath, SaveFormat.Docx);
        }

        // Example usage.
        public static void Main()
        {
            string templatePath = @"C:\Templates\SampleTemplate.docm";
            string outputPath = @"C:\Output\ResultDocument.docx";

            string html = @"
                <h1 style='color:#2E86C1;'>Dynamic Title</h1>
                <p style='font-size:12pt;'>This paragraph is inserted from HTML.</p>";

            // Insert using source (HTML) styles.
            InsertHtmlIntoTemplate(templatePath, outputPath, html, useSourceStyles: true);
        }
    }
}
