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
        // Create a sample DOCX document with some text and an image.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample DOCX document that will be converted to MHTML.");
        
        // Create a simple PNG image using Aspose.Drawing and save it to a file.
        const string imagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Insert the created image into the document.
        builder.InsertImage(imagePath);

        // Save the document as DOCX (input file for conversion).
        const string inputDocxPath = "input.docx";
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputDocxPath);

        // Configure save options for MHTML with embedded fonts and CID URLs for resources.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportFontResources = true,
            ExportCidUrlsForMhtmlResources = true
        };

        // Convert and save to MHTML.
        const string outputMhtmlPath = "output.mht";
        doc.Save(outputMhtmlPath, saveOptions);

        // Validate that the MHTML file was created.
        if (!File.Exists(outputMhtmlPath))
            throw new InvalidOperationException("The MHTML output file was not created.");

        // Clean up temporary files (optional).
        File.Delete(imagePath);
        File.Delete(inputDocxPath);
    }
}
