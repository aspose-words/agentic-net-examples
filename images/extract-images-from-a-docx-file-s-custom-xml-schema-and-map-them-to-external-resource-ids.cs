using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // 2. Create a new document and insert the image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape insertedShape = builder.InsertImage(sampleImagePath);

        // 3. Add a custom XML part that maps an external resource ID to the image index
        string customXml = @"
<Resources>
    <Resource>
        <Id>res1</Id>
        <ImageIndex>0</ImageIndex>
    </Resource>
</Resources>";
        // Add the XML part with a generated ID
        doc.CustomXmlParts.Add(Guid.NewGuid().ToString(), customXml);

        // 4. Save the document
        string docPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docPath);

        // 5. Load the document for extraction
        Document loadedDoc = new Document(docPath);

        // Retrieve the custom XML part (first one)
        if (loadedDoc.CustomXmlParts.Count == 0)
            throw new InvalidOperationException("No custom XML parts found in the document.");

        // Data property returns a byte[]; convert to string
        string xmlData = Encoding.UTF8.GetString(loadedDoc.CustomXmlParts[0].Data);
        XDocument xDoc = XDocument.Parse(xmlData);

        // 6. Collect all shapes that contain images
        var imageShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .Where(s => s.HasImage)
                                   .ToList();

        if (imageShapes.Count == 0)
            throw new InvalidOperationException("No images found in the document.");

        // 7. Map each resource ID to its corresponding image and save the image file
        var resources = xDoc.Descendants("Resource");
        int extractedCount = 0;
        foreach (var res in resources)
        {
            string resourceId = res.Element("Id")?.Value;
            string indexStr = res.Element("ImageIndex")?.Value;
            if (string.IsNullOrEmpty(resourceId) || string.IsNullOrEmpty(indexStr))
                continue;

            if (!int.TryParse(indexStr, out int imgIndex) || imgIndex < 0 || imgIndex >= imageShapes.Count)
                continue;

            Shape shape = imageShapes[imgIndex];
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outImagePath = Path.Combine(outputDir, $"{resourceId}{extension}");
            shape.ImageData.Save(outImagePath);
            extractedCount++;
        }

        // 8. Validation
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted based on the custom XML mapping.");

        Console.WriteLine($"Extraction completed. {extractedCount} image(s) saved to '{outputDir}'.");
    }
}
