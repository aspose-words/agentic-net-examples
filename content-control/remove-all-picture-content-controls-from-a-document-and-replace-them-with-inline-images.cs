using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample document that contains a picture content control.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a block-level picture content control and add it to the document body.
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Block);
        pictureSdt.Title = "SamplePicture";
        pictureSdt.Tag = "sample-picture";

        // Append the content control to the body.
        doc.FirstSection.Body.AppendChild(pictureSdt);

        // Inside the content control add a paragraph that will hold the image.
        Paragraph sdtParagraph = new Paragraph(doc);
        pictureSdt.AppendChild(sdtParagraph);

        // ---------------------------------------------------------------
        // Generate a simple red square image using Aspose.Drawing.
        // ---------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Fill the bitmap with red pixels.
            for (int y = 0; y < bitmap.Height; y++)
            {
                for (int x = 0; x < bitmap.Width; x++)
                {
                    bitmap.SetPixel(x, y, Color.Red);
                }
            }

            // Save the bitmap to a memory stream in PNG format.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the picture content control.
                DocumentBuilder sdtBuilder = new DocumentBuilder(doc);
                sdtBuilder.MoveTo(sdtParagraph);
                sdtBuilder.InsertImage(imageStream);
            }
        }

        // Save the seed document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // ---------------------------------------------------------------
        // Load the document and replace picture content controls with inline images.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Find all picture content controls in the document.
        var pictureSdts = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(s => s.SdtType == SdtType.Picture)
            .ToList();

        foreach (StructuredDocumentTag sdt in pictureSdts)
        {
            // Locate the first shape (image) inside the content control.
            Shape? shape = sdt.GetChildNodes(NodeType.Shape, true)
                .OfType<Shape>()
                .FirstOrDefault();

            if (shape != null && shape.HasImage)
            {
                // Extract the image bytes.
                byte[] imageBytes = shape.ImageData.ImageBytes;

                // Insert the image inline at the position of the content control.
                DocumentBuilder replaceBuilder = new DocumentBuilder(loadedDoc);
                replaceBuilder.MoveTo(sdt);
                using (MemoryStream imgStream = new MemoryStream(imageBytes))
                {
                    replaceBuilder.InsertImage(imgStream);
                }

                // Remove the original picture content control.
                sdt.Remove();
            }
        }

        // Save the resulting document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
