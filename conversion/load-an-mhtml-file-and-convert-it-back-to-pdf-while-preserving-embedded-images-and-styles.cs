using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string mhtmlPath = "sample.mhtml";
        string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample Word document with styled text and an image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply some styling to the text.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;
        builder.Writeln("This is a styled paragraph with an embedded image:");

        // Create a simple bitmap using Aspose.Drawing (no System.Drawing usage).
        Bitmap bitmap = new Bitmap(100, 100);
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill the bitmap with a solid blue color.
            graphics.Clear(Color.Blue);
        }

        // Save the bitmap to a memory stream in PNG format.
        using (MemoryStream imageStream = new MemoryStream())
        {
            bitmap.Save(imageStream, ImageFormat.Png);
            imageStream.Position = 0; // Reset position before reading.

            // Insert the image into the document.
            builder.InsertImage(imageStream);
        }

        // Dispose of the bitmap now that it is no longer needed.
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Save the document as MHTML (Web archive) to preserve images/styles.
        // -----------------------------------------------------------------
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath))
            throw new InvalidOperationException("MHTML file was not created.");

        // -----------------------------------------------------------------
        // 3. Load the MHTML file back into a new Document instance.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(mhtmlPath);

        // -----------------------------------------------------------------
        // 4. Convert the loaded document to PDF, preserving embedded content.
        // -----------------------------------------------------------------
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // The example finishes here. No interactive prompts are used.
    }
}
