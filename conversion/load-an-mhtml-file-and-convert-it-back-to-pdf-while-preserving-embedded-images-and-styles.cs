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
        // Paths for temporary files
        const string imagePath = "sample.png";
        const string mhtmlPath = "sample.mhtml";
        const string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // Create a simple image using Aspose.Drawing (no System.Drawing usage)
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // ---------------------------------------------------------------
        // Build a Word document, insert text and the created image, save as MHTML
        // ---------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample MHTML document with an embedded image and style.");
        builder.InsertImage(imagePath);
        sourceDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // ---------------------------------------------------------------
        // Load the MHTML file and convert it to PDF while preserving content
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(mhtmlPath);
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // ---------------------------------------------------------------
        // Validation: ensure the PDF file was created
        // ---------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF conversion failed; output file was not created.");

        // Clean up temporary files (optional)
        File.Delete(imagePath);
        File.Delete(mhtmlPath);
    }
}
