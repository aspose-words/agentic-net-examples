using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for input and output files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample GIF image (static for simplicity).
        // -----------------------------------------------------------------
        string gifPath = Path.Combine(workDir, "sample.gif");
        CreateSampleGif(gifPath);

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the GIF image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(workDir, "DocumentWithGif.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all GIF images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int gifIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Process only GIF images.
            if (shape.ImageData.ImageType != ImageType.Gif) continue;

            // Save the extracted GIF.
            string extractedGif = Path.Combine(workDir, $"extracted_{gifIndex}.gif");
            shape.ImageData.Save(extractedGif);
            if (!File.Exists(extractedGif))
                throw new InvalidOperationException($"Failed to save extracted GIF: {extractedGif}");

            // -----------------------------------------------------------------
            // 4. Convert the extracted GIF to an MP4 video clip.
            //    (Placeholder conversion – copies the GIF to an MP4 file.
            //     Real conversion would require a video processing library such as FFmpeg.)
            // -----------------------------------------------------------------
            string mp4Path = Path.Combine(workDir, $"video_{gifIndex}.mp4");
            File.Copy(extractedGif, mp4Path, overwrite: true);
            if (!File.Exists(mp4Path))
                throw new InvalidOperationException($"Failed to create MP4 placeholder: {mp4Path}");

            Console.WriteLine($"GIF extracted to: {extractedGif}");
            Console.WriteLine($"MP4 placeholder created at: {mp4Path}");
            gifIndex++;
        }

        // Validate that at least one GIF was processed.
        if (gifIndex == 0)
            throw new InvalidOperationException("No GIF images were found in the document.");

        // Cleanup (optional): uncomment the following line to delete the working directory after execution.
        // Directory.Delete(workDir, recursive: true);
    }

    // Helper method to create a simple static GIF image using Aspose.Drawing.
    private static void CreateSampleGif(string filePath)
    {
        const int width = 200;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            graphics.DrawString(
                "Sample GIF",
                new Aspose.Drawing.Font("Arial", 20),
                Aspose.Drawing.Brushes.Black,
                new Aspose.Drawing.PointF(10, 40));

            // Save as GIF.
            bitmap.Save(filePath, ImageFormat.Gif);
        }
    }
}
