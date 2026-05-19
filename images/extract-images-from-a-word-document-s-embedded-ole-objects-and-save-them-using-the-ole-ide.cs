using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json; // Included as required package

public class ExtractOleImages
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create a sample icon image ----------
        string iconPath = Path.Combine(artifactsDir, "icon.png");
        using (Bitmap bitmap = new Bitmap(64, 64))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightGray);
                // Draw a simple rectangle as visual content
                g.DrawRectangle(new Pen(Aspose.Drawing.Color.Blue, 2), 8, 8, 48, 48);
            }
            bitmap.Save(iconPath, ImageFormat.Png);
        }

        // ---------- Create a sample OLE data stream ----------
        byte[] oleContent = Encoding.UTF8.GetBytes("Sample OLE embedded content");
        using (MemoryStream oleStream = new MemoryStream(oleContent))
        {
            // Reset position before use
            oleStream.Position = 0;

            // ---------- Build a document with an OLE object ----------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Document containing an OLE object with an icon:");

            // Load the icon into a stream
            using (MemoryStream iconStream = new MemoryStream(File.ReadAllBytes(iconPath)))
            {
                iconStream.Position = 0;
                // Insert the OLE object as an icon (ProgId "Package" works for generic data)
                builder.InsertOleObject(oleStream, "Package", true, iconStream);
            }

            // Save the document
            string docPath = Path.Combine(artifactsDir, "OleDocument.docx");
            doc.Save(docPath);
        }

        // ---------- Load the document and extract OLE icon images ----------
        Document loadedDoc = new Document(Path.Combine(artifactsDir, "OleDocument.docx"));
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Process only OLE object shapes
            if (shape.ShapeType == ShapeType.OleObject)
            {
                OleFormat oleFormat = shape.OleFormat;
                string progId = !string.IsNullOrEmpty(oleFormat?.ProgId) ? oleFormat.ProgId : "OleObject";

                // If the OLE shape has an associated image (icon), save it
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outFileName = $"{progId}_{imageIndex}{extension}";
                    string outPath = Path.Combine(artifactsDir, outFileName);
                    shape.ImageData.Save(outPath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (imageIndex == 0)
            throw new InvalidOperationException("No OLE icon images were extracted.");

        // Optional: list extracted files (non‑interactive)
        Console.WriteLine($"Extracted {imageIndex} OLE icon image(s) to folder: {artifactsDir}");
    }
}
