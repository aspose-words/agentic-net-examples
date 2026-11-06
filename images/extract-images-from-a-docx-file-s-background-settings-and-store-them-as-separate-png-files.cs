using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic background image (bg.png).
        // -----------------------------------------------------------------
        const string backgroundImagePath = "bg.png";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 50, 50, 100, 100);
                }
            }
            bitmap.Save(backgroundImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Create a DOCX document and insert the background image so that
        //    it covers the printable page area.
        // -----------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Calculate printable page dimensions.
        double pageWidth = doc.FirstSection.PageSetup.PageWidth -
                           doc.FirstSection.PageSetup.LeftMargin -
                           doc.FirstSection.PageSetup.RightMargin;
        double pageHeight = doc.FirstSection.PageSetup.PageHeight -
                            doc.FirstSection.PageSetup.TopMargin -
                            doc.FirstSection.PageSetup.BottomMargin;

        // Insert the image using InsertImage (preferred API). This returns the
        // underlying Shape which we can further configure.
        Shape bgShape = builder.InsertImage(backgroundImagePath, (float)pageWidth, (float)pageHeight);
        bgShape.WrapType = WrapType.None;
        bgShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        bgShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        bgShape.BehindText = true;

        // Save the document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images (including the background).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outputImagePath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(outputImagePath);
                if (!File.Exists(outputImagePath))
                    throw new Exception($"Failed to save extracted image to '{outputImagePath}'.");
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // (Optional) Clean up temporary files.
        // -----------------------------------------------------------------
        // File.Delete(backgroundImagePath);
        // File.Delete(docPath);
    }
}
