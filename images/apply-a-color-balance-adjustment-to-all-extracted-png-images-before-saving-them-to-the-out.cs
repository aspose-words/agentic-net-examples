using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string inputImagePath = Path.Combine(artifactsDir, "input.png");
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill with a simple gradient background.
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(
                    Aspose.Drawing.Color.FromArgb(255, 100, 150, 200)))
                {
                    g.FillRectangle(brush, 0, 0, bitmap.Width, bitmap.Height);
                }
            }
            bitmap.Save(inputImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Insert the PNG into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // -----------------------------------------------------------------
            // 4. Save the shape's image to a memory stream.
            // -----------------------------------------------------------------
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // -----------------------------------------------------------------
                // 5. Load the image with Aspose.Drawing.
                // -----------------------------------------------------------------
                using (Aspose.Drawing.Image original = Aspose.Drawing.Image.FromStream(imageStream))
                {
                    // -----------------------------------------------------------------
                    // 6. Apply a simple color‑balance adjustment.
                    //    Increase Red, keep Green, decrease Blue.
                    // -----------------------------------------------------------------
                    using (Aspose.Drawing.Bitmap adjusted = new Aspose.Drawing.Bitmap(original.Width, original.Height))
                    {
                        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(adjusted))
                        {
                            float redFactor = 1.2f;
                            float greenFactor = 1.0f;
                            float blueFactor = 0.8f;

                            float[][] matrixElements =
                            {
                                new float[] { redFactor, 0, 0, 0, 0 },
                                new float[] { 0, greenFactor, 0, 0, 0 },
                                new float[] { 0, 0, blueFactor, 0, 0 },
                                new float[] { 0, 0, 0, 1, 0 },
                                new float[] { 0, 0, 0, 0, 1 }
                            };
                            Aspose.Drawing.Imaging.ColorMatrix colorMatrix = new Aspose.Drawing.Imaging.ColorMatrix(matrixElements);
                            Aspose.Drawing.Imaging.ImageAttributes imgAttr = new Aspose.Drawing.Imaging.ImageAttributes();
                            imgAttr.SetColorMatrix(colorMatrix);

                            graphics.DrawImage(
                                original,
                                new Rectangle(0, 0, adjusted.Width, adjusted.Height),
                                0,
                                0,
                                original.Width,
                                original.Height,
                                GraphicsUnit.Pixel,
                                imgAttr);
                        }

                        // -----------------------------------------------------------------
                        // 7. Save the adjusted PNG to the output folder.
                        // -----------------------------------------------------------------
                        string outputPath = Path.Combine(artifactsDir, $"extracted_{imageIndex}.png");
                        adjusted.Save(outputPath, Aspose.Drawing.Imaging.ImageFormat.Png);

                        // Validate that the file was created.
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to save adjusted image to '{outputPath}'.");

                        imageIndex++;
                    }
                }
            }
        }

        // Ensure at least one image was extracted and saved.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
