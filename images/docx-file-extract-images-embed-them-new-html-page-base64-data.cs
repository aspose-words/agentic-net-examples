using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractImagesToBase64Html
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Add a simple paragraph so the document is not completely empty.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document generated at runtime.");

        // Configure HTML save options to embed images as Base64 data.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportImagesAsBase64 = true,
            PrettyFormat = true
        };

        // Save the document to a memory stream using the configured options.
        using (MemoryStream htmlStream = new MemoryStream())
        {
            doc.Save(htmlStream, htmlOptions);

            // Convert the stream contents to a UTF‑8 string containing the HTML.
            string htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());

            // Write the resulting HTML to a file in the current directory.
            string outputHtml = Path.Combine(Environment.CurrentDirectory, "DocumentWithEmbeddedImages.html");
            File.WriteAllText(outputHtml, htmlContent, Encoding.UTF8);

            Console.WriteLine($"HTML saved to: {outputHtml}");
        }
    }
}
