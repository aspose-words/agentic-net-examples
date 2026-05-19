using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class HtmlToPdfConverter
{
    public static void Main()
    {
        // Create a simple HTML content locally.
        const string htmlContent = "<html><body><h1>Sample Heading</h1><p>This is a sample HTML page.</p></body></html>";

        // Convert the HTML string to a byte array.
        byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);

        // Load the HTML from a memory stream using HtmlLoadOptions.
        using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
        {
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                LoadFormat = LoadFormat.Html
            };

            Document document = new Document(htmlStream, loadOptions);

            // Define the output PDF file path.
            const string outputPath = "output.pdf";

            // Save the document as PDF.
            document.Save(outputPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The PDF file was not created.");
        }
    }
}
