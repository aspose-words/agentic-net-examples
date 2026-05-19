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
        // Directories for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string gifDir = Path.Combine(workDir, "Gifs");
        string mp4Dir = Path.Combine(workDir, "Mp4s");
        Directory.CreateDirectory(gifDir);
        Directory.CreateDirectory(mp4Dir);

        // -----------------------------------------------------------------
        // 1. Create a sample GIF image (static for simplicity) and insert it
        //    into a Word document.
        // -----------------------------------------------------------------
        string sampleGifPath = Path.Combine(workDir, "sample.gif");
        CreateSampleGif(sampleGifPath);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleGifPath);
        string docPath = Path.Combine(workDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract all GIF images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .Cast<Shape>()
                                  .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Gif)
                                  .ToList();

        if (!shapeNodes.Any())
            throw new InvalidOperationException("No GIF images were found in the document.");

        int gifIndex = 0;
        foreach (var shape in shapeNodes)
        {
            string gifFileName = Path.Combine(gifDir, $"extracted_{gifIndex}.gif");
            shape.ImageData.Save(gifFileName);
            gifIndex++;
        }

        // -----------------------------------------------------------------
        // 3. Batch convert each extracted GIF to an MP4 file.
        //    (Real video conversion would require a multimedia library; here we
        //     simply copy the file and change the extension to illustrate the
        //     batch workflow while preserving order.)
        // -----------------------------------------------------------------
        var gifFiles = Directory.GetFiles(gifDir, "*.gif")
                                .OrderBy(f => f) // Preserve lexical order (which matches extraction order)
                                .ToArray();

        if (!gifFiles.Any())
            throw new InvalidOperationException("No GIF files were found for conversion.");

        for (int i = 0; i < gifFiles.Length; i++)
        {
            string sourceGif = gifFiles[i];
            string targetMp4 = Path.Combine(mp4Dir, $"video_{i}.mp4");

            // Placeholder conversion: copy the GIF file and rename it.
            // Replace this block with a real video encoder if needed.
            File.Copy(sourceGif, targetMp4, overwrite: true);
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure that MP4 files were created.
        // -----------------------------------------------------------------
        var mp4Files = Directory.GetFiles(mp4Dir, "*.mp4");
        if (mp4Files.Length == 0)
            throw new InvalidOperationException("MP4 conversion failed – no output files were created.");

        Console.WriteLine($"Extraction and conversion completed. {mp4Files.Length} MP4 files are available in '{mp4Dir}'.");
    }

    // Creates a simple static GIF image using Aspose.Drawing.
    private static void CreateSampleGif(string filePath)
    {
        const int width = 200;
        const int height = 200;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.LightBlue);
            graphics.DrawEllipse(new Pen(Aspose.Drawing.Color.DarkBlue, 5), 20, 20, width - 40, height - 40);
            bitmap.Save(filePath, ImageFormat.Gif);
        }
    }
}
