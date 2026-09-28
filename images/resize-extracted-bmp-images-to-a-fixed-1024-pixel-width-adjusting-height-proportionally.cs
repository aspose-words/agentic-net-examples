using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample BMP image.
        const string sampleBmpPath = "sample.bmp";
        const int sampleWidth = 200;
        const int sampleHeight = 100;
        using (Bitmap bmp = new Bitmap(sampleWidth, sampleHeight))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.Red), 0, 0, sampleWidth, sampleHeight);
            }
            bmp.Save(sampleBmpPath);
        }

        // Create a Word document and insert the BMP image.
        const string docPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        doc.Save(docPath);

        // Load the document for image extraction.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the extracted BMP.
                using (Bitmap originalBmp = new Bitmap(imageStream))
                {
                    // Calculate new dimensions (width = 1024, height proportional).
                    const int targetWidth = 1024;
                    int originalWidth = originalBmp.Width;
                    int originalHeight = originalBmp.Height;
                    int targetHeight = (int)Math.Round((double)originalHeight * targetWidth / originalWidth);

                    // Resize the bitmap.
                    using (Bitmap resizedBmp = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBmp))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(originalBmp, 0, 0, targetWidth, targetHeight);
                        }

                        // Save the resized image.
                        string outputPath = $"resized-{imageIndex}.bmp";
                        resizedBmp.Save(outputPath);

                        // Validate that the file was created.
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to create resized image: {outputPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found to resize.");
    }
}
