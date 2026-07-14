using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image.
        const string sampleImagePath = "sample.png";
        Bitmap bitmap = new Bitmap(200, 200);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        graphics.FillEllipse(
            new SolidBrush(Color.Blue),
            25, 25, 150, 150);
        bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();

        // Step 2: Create a Word document and insert the sample image.
        const string wordPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(wordPath);

        // Step 3: Load the Word document and extract all images.
        Document loadedDoc = new Document(wordPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedImagePaths = new List<string>();
        int imageIndex = 1;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedPath = $"extracted-{imageIndex}.png";
                shape.ImageData.Save(extractedPath);
                extractedImagePaths.Add(extractedPath);
                imageIndex++;
            }
        }

        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the Word document.");

        // Validation: ensure output files exist.
        if (!File.Exists(wordPath))
            throw new FileNotFoundException("Word document was not created.", wordPath);

        foreach (string imgPath in extractedImagePaths)
        {
            if (!File.Exists(imgPath))
                throw new FileNotFoundException("Extracted image file missing.", imgPath);
        }

        // Note: Embedding images into a PowerPoint presentation would require Aspose.Slides,
        // which is not part of the allowed package set for this example. The extracted images
        // are saved to disk and can be used in a presentation with appropriate tooling.
    }
}
