using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare the output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create sample files needed for the OLE object.
        // -----------------------------------------------------------------
        // Text file that will be embedded as an OLE package.
        string sampleTextPath = Path.Combine(artifactsDir, "sample.txt");
        File.WriteAllText(sampleTextPath, "This is a sample text file used as OLE data.");

        // Create a simple 32x32 PNG icon that will be used as the visual representation.
        string iconPath = Path.Combine(artifactsDir, "icon.png");
        using (Bitmap bitmap = new Bitmap(32, 32))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                g.DrawRectangle(new Pen(Aspose.Drawing.Color.DarkBlue, 2), 4, 4, 24, 24);
            }
            bitmap.Save(iconPath);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document that contains an embedded OLE object
        //    displayed as an icon (the image we just created).
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "OleDocument.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object as an icon using the overload that accepts an image stream.
        using (FileStream iconStream = File.OpenRead(iconPath))
        {
            // Parameters: fileName, isLinked, asIcon, presentation (image stream)
            builder.InsertOleObject(sampleTextPath, false, true, iconStream);
        }

        // Save the document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract the icon images from OLE objects.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // We are interested only in OLE objects that have an image (icon).
            if (shape.ShapeType == ShapeType.OleObject && shape.HasImage)
            {
                OleFormat ole = shape.OleFormat;
                string progId = ole?.ProgId ?? "OleObject";

                // Determine the proper file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputFileName = $"{progId}_{extractedCount}{extension}";
                string outputPath = Path.Combine(artifactsDir, outputFileName);

                // Save the image.
                shape.ImageData.Save(outputPath);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No OLE object images were extracted.");
    }
}
