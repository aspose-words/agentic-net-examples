using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static async Task Main()
    {
        // Path to the DOCX template.
        string templatePath = "Template.docx";

        // Path for the generated PDF.
        string outputPdfPath = "Result.pdf";

        // Load the template document.
        Document doc = new Document(templatePath); // uses Document(string) constructor

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first image from a REST endpoint at a bookmark named "Image1".
        await InsertImageFromUrlAsync(builder, "https://example.com/image1.jpg", "Image1");

        // Insert the second image at the end of the document (no bookmark).
        await InsertImageFromUrlAsync(builder, "https://example.com/image2.png", null);

        // Save the modified document as PDF. The format is inferred from the file extension.
        doc.Save(outputPdfPath); // uses Document.Save(string)
    }

    // Downloads an image from a URL and inserts it into the document.
    // If bookmarkName is provided, the cursor is moved to that bookmark before insertion.
    static async Task InsertImageFromUrlAsync(DocumentBuilder builder, string imageUrl, string bookmarkName)
    {
        using HttpClient httpClient = new HttpClient();
        byte[] imageBytes = await httpClient.GetByteArrayAsync(imageUrl);

        if (!string.IsNullOrEmpty(bookmarkName))
        {
            builder.MoveToBookmark(bookmarkName);
        }

        // Insert the image from the downloaded byte array (inline, 100% scale).
        builder.InsertImage(imageBytes);
    }
}
