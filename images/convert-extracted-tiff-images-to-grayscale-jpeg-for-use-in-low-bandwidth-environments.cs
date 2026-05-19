using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace AsposeWordsImageConversion
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a folder for all generated files.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // 1. Create a deterministic sample TIFF image.
            string tiffPath = Path.Combine(artifactsDir, "sample.tif");
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.LightBlue);
                    g.FillRectangle(new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkRed), 50, 50, 100, 100);
                }
                bitmap.Save(tiffPath);
            }

            // 2. Insert the TIFF image into a new Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(tiffPath);
            string docPath = Path.Combine(artifactsDir, "DocumentWithTiff.docx");
            doc.Save(docPath);

            // 3. Load the document (demonstrating the load lifecycle).
            Document loadedDoc = new Document(docPath);

            // 4. Extract every image, force grayscale, and save it as JPEG.
            NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Force the image to display in grayscale.
                shape.ImageData.GrayScale = true;

                // Save the image as JPEG with a deterministic name.
                string jpegPath = Path.Combine(artifactsDir, $"ExtractedImage_{imageIndex}.jpg");
                shape.ImageData.Save(jpegPath);
                imageIndex++;
            }

            // 5. Validate that at least one JPEG was created.
            if (imageIndex == 0)
                throw new InvalidOperationException("No images were found and converted.");

            // Program ends – all files are written to the Artifacts folder.
        }
    }
}
