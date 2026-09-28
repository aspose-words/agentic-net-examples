using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample document with styled text and an embedded image.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add a heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Heading");

        // Create a simple bitmap image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with blue using Aspose.Drawing.Color.
                graphics.Clear(Aspose.Drawing.Color.Blue);
            }

            using (MemoryStream imageStream = new MemoryStream())
            {
                // Save the bitmap to a memory stream in PNG format.
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Add a styled paragraph (bold text). Font color is left as default to avoid System.Drawing usage.
        builder.Font.Bold = true;
        builder.Writeln("Styled text with bold.");

        // Save the document as MHTML.
        string mhtmlPath = "sample.mht";
        sourceDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Load the MHTML file.
        Document loadedDoc = new Document(mhtmlPath);

        // Convert the loaded document to PDF.
        string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
