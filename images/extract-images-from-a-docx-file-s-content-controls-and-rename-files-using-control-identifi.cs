using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic sample image (100x100 white PNG).
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // 2. Build a DOCX with two content controls, each containing an image.
        string docPath = Path.Combine(outputDir, "sample.docx");
        CreateDocumentWithContentControls(docPath, sampleImagePath);

        // 3. Load the document and extract images from content controls.
        Document doc = new Document(docPath);
        int extractedCount = 0;

        // Get all StructuredDocumentTag (content control) nodes.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes.OfType<StructuredDocumentTag>())
        {
            // Determine a base name for extracted files using the control's Title (if set) or Id.
            string baseName = !string.IsNullOrEmpty(sdt.Title) ? sdt.Title : $"SDT_{sdt.Id}";

            // Find all Shape nodes inside the current content control.
            NodeCollection shapeNodes = sdt.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{baseName}{extension}";
                    string imagePath = Path.Combine(outputDir, imageFileName);

                    // Save the image.
                    shape.ImageData.Save(imagePath);
                    extractedCount++;

                    // Validate that the file was created.
                    if (!File.Exists(imagePath))
                        throw new InvalidOperationException($"Failed to save extracted image: {imagePath}");
                }
            }
        }

        // Ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from content controls.");

        // Indicate success (no interactive I/O required).
        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to \"{outputDir}\".");
    }

    // Creates a simple white PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create bitmap.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        try
        {
            // Fill with white.
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                graphics.Clear(Aspose.Drawing.Color.White);
            }
            finally
            {
                graphics.Dispose();
            }

            // Save as PNG.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }

        // Validate that the image file exists.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at \"{filePath}\".");
    }

    // Creates a DOCX containing two content controls, each with an image.
    private static void CreateDocumentWithContentControls(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First content control.
        StructuredDocumentTag sdt1 = new StructuredDocumentTag(builder.Document, SdtType.RichText, MarkupLevel.Block);
        sdt1.Title = "ControlOne";

        // Append the SDT to the document body (valid insertion point for a block node).
        doc.FirstSection.Body.AppendChild(sdt1);

        // Ensure the SDT contains a paragraph to host the image.
        Paragraph para1 = new Paragraph(builder.Document);
        sdt1.AppendChild(para1);
        builder.MoveTo(para1);
        builder.InsertImage(imagePath);

        // Add a page break between controls.
        builder.InsertBreak(BreakType.PageBreak);

        // Second content control.
        StructuredDocumentTag sdt2 = new StructuredDocumentTag(builder.Document, SdtType.RichText, MarkupLevel.Block);
        sdt2.Title = "ControlTwo";

        doc.FirstSection.Body.AppendChild(sdt2);

        Paragraph para2 = new Paragraph(builder.Document);
        sdt2.AppendChild(para2);
        builder.MoveTo(para2);
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);

        // Validate that the document file exists.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save document at \"{docPath}\".");
    }
}
