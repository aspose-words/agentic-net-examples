using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the first page has a different header/footer.
        doc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Base64‑encoded PNG image (a simple 16×16 blue square).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAFklEQVR42mNk+M9Qz0AEYBxVSFIAABcAAf+Z6+UAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Move to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Center the image horizontally.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert the logo image from the memory stream.
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            builder.InsertImage(ms);
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstPageHeader.docx");
        doc.Save(outputPath);
    }
}
