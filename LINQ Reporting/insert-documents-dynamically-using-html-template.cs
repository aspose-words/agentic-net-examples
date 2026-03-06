using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDynamicInsert
{
    class Program
    {
        static void Main()
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the newly created document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Load the HTML template from a file. The template can contain placeholders
            // that will be replaced by the application before insertion if needed.
            string htmlTemplatePath = "Template.html";
            string htmlContent = File.ReadAllText(htmlTemplatePath);

            // Insert the HTML content into the document.
            // Use builder formatting as the base and remove the extra empty paragraph
            // that InsertHtml normally adds after a block‑level element.
            builder.InsertHtml(htmlContent,
                HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

            // Save the resulting document to disk.
            string outputPath = "Result.docx";
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
