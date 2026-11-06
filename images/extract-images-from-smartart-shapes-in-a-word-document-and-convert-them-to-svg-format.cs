using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // ---------- Create a deterministic sample PNG image ----------
        const string sampleImagePath = "sample.png";
        const int imgWidth = 100;
        const int imgHeight = 100;

        // Create bitmap and draw a simple rectangle.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath);
        }

        // Verify that the sample image was created.
        if (!File.Exists(sampleImagePath))
            throw new Exception($"Failed to create sample image: {sampleImagePath}");

        // ---------- Create a Word document and insert the image ----------
        const string docPath = "SmartArtDoc.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as a picture shape.
        Shape pictureShape = builder.InsertImage(sampleImagePath);
        // Ensure the shape is added to the document.
        if (pictureShape == null)
            throw new Exception("Failed to insert image into the document.");

        // Save the document.
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document: {docPath}");

        // ---------- Load the document and extract images ----------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Save image bytes to a memory stream.
                byte[] imageBytes;
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0;
                    imageBytes = ms.ToArray();
                }

                // Convert the image to an SVG that embeds the bitmap as base64.
                string base64 = Convert.ToBase64String(imageBytes);
                string svgContent = $"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{imgWidth}\" height=\"{imgHeight}\">" +
                                    $"<image href=\"data:image/png;base64,{base64}\" width=\"{imgWidth}\" height=\"{imgHeight}\"/>" +
                                    $"</svg>";

                string svgPath = $"extracted-image-{imageIndex}.svg";
                File.WriteAllText(svgPath, svgContent, Encoding.UTF8);

                if (!File.Exists(svgPath))
                    throw new Exception($"Failed to create SVG file: {svgPath}");

                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new Exception("No images were extracted from the document.");

        // Optional cleanup (uncomment if desired).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
