using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a sample cover‑art image (200×200 white background).
        // -----------------------------------------------------------------
        string coverPath = Path.Combine(workDir, "cover.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a simple red ellipse to make the image recognizable.
            using (Pen pen = new Pen(Color.Red, 5))
            {
                g.DrawEllipse(pen, 20, 20, 160, 160);
            }
            bitmap.Save(coverPath);
        }

        // -----------------------------------------------------------------
        // 2. Create a DOCX file and insert the image as if it were audio cover art.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(workDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image inline; in a real scenario this would be the audio cover art.
        builder.InsertImage(coverPath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images from shapes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Obtain the raw image bytes from the shape.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into an Aspose.Drawing.Image.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Image img = Image.FromStream(ms))
            {
                // Save the image as JPEG regardless of its original format.
                string outFile = Path.Combine(workDir, $"extracted_{imageIndex}.jpg");
                img.Save(outFile, ImageFormat.Jpeg);
                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // -----------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // The program finishes automatically; all files are written to the Work folder.
    }
}
