using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (PNG)
        const string inputImagePath = "input.png";
        const int imgWidth = 200;
        const int imgHeight = 100;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(inputImagePath);
        }

        // Create a Word document and insert the image (simulating a SmartArt shape)
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Reload the document to simulate a real scenario
        Document loadedDoc = new Document(docPath);

        // Extract images from shapes (including those that could be SmartArt)
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(Aspose.Words.NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;
                    byte[] imgBytes = imgStream.ToArray();
                    string base64 = Convert.ToBase64String(imgBytes);

                    // Determine image format for data URI (use PNG as we inserted PNG)
                    const string imageMime = "image/png";

                    // Build simple SVG containing the image as base64
                    double widthPt = shape.Width;   // points
                    double heightPt = shape.Height; // points
                    // Convert points to pixels (1 point = 1/72 inch, assume 96 DPI)
                    const double dpi = 96.0;
                    double widthPx = widthPt * dpi / 72.0;
                    double heightPx = heightPt * dpi / 72.0;

                    string svgContent = $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<svg xmlns=""http://www.w3.org/2000/svg"" width=""{widthPx:F2}px"" height=""{heightPx:F2}px"">
  <image href=""data:{imageMime};base64,{base64}"" width=""{widthPx:F2}"" height=""{heightPx:F2}""/>
</svg>";

                    string svgPath = $"extracted-image-{extractedCount + 1}.svg";
                    File.WriteAllText(svgPath, svgContent, Encoding.UTF8);
                    extractedCount++;
                }
            }
        }

        // Validation
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional cleanup (commented out)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
