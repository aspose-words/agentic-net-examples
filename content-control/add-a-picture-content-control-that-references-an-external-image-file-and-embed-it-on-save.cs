using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph before the picture content control.
        builder.Writeln("Below is a picture content control with an embedded image:");

        // Insert a picture content control at the current cursor position.
        StructuredDocumentTag pictureSdt = builder.InsertStructuredDocumentTag(SdtType.Picture);
        pictureSdt.Title = "SamplePicture";
        pictureSdt.Tag = "sample-pic";

        // Ensure a sample image file exists in the working directory.
        const string imageFileName = "sample.png";
        if (!File.Exists(imageFileName))
        {
            // A 1x1 pixel transparent PNG (base64 encoded).
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imageFileName, pngBytes);
        }

        // Insert the image inside the picture content control.
        builder.InsertImage(imageFileName);

        // Save the document; the image is embedded automatically.
        doc.Save("PictureContentControl.docx");
    }
}
