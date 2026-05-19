using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample BMP image (800x800, solid blue) and save it.
        // -----------------------------------------------------------------
        string inputBmpPath = Path.Combine(artifactsDir, "input.bmp");
        using (Bitmap bmp = new Bitmap(800, 800))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Blue);
            }
            bmp.Save(inputBmpPath, ImageFormat.Bmp);
        }

        // ---------------------------------------------------------------
        // 2. Create a Word document and insert the BMP image.
        // ---------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputBmpPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithBmp.docx");
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 3. Load the document and process each BMP image.
        //    Resize the image until its file size is <= 200 KB.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Bmp) continue;

            // Extract the original image into a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the bitmap from the stream.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Quick check: if the original BMP is already small enough, keep it.
                    using (MemoryStream sizeCheck = new MemoryStream())
                    {
                        originalBitmap.Save(sizeCheck, ImageFormat.Bmp);
                        if (sizeCheck.Length <= 200 * 1024)
                        {
                            // No resizing needed.
                            continue;
                        }
                    }

                    // Iteratively scale down the bitmap by 10% until size <= 200 KB.
                    const double scaleFactor = 0.9;
                    Bitmap workingBitmap = originalBitmap;
                    MemoryStream resizedStream = null;

                    while (true)
                    {
                        int newWidth = (int)(workingBitmap.Width * scaleFactor);
                        int newHeight = (int)(workingBitmap.Height * scaleFactor);
                        if (newWidth < 1 || newHeight < 1) break; // Prevent zero size.

                        Bitmap scaledBitmap = new Bitmap(newWidth, newHeight);
                        using (Graphics g = Graphics.FromImage(scaledBitmap))
                        {
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(workingBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Dispose previous bitmap if it was created in a prior loop.
                        if (!ReferenceEquals(workingBitmap, originalBitmap))
                            workingBitmap.Dispose();

                        workingBitmap = scaledBitmap;

                        // Save to a stream and test the size.
                        resizedStream?.Dispose();
                        resizedStream = new MemoryStream();
                        workingBitmap.Save(resizedStream, ImageFormat.Bmp);

                        if (resizedStream.Length <= 200 * 1024)
                            break; // Desired size reached.
                    }

                    // Replace the shape's image with the resized bitmap.
                    resizedStream.Position = 0;
                    shape.ImageData.SetImage(resizedStream);

                    // Clean up.
                    resizedStream.Dispose();
                    if (!ReferenceEquals(workingBitmap, originalBitmap))
                        workingBitmap.Dispose();
                }
            }
        }

        // ---------------------------------------------------------------
        // 4. Save the document with resized BMP images.
        // ---------------------------------------------------------------
        string outputDocPath = Path.Combine(artifactsDir, "DocumentResizedBmp.docx");
        loadedDoc.Save(outputDocPath);

        // Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new Exception("The output document was not created.");
    }
}
