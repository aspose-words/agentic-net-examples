using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Drawing; // Aspose.Drawing namespace for Bitmap, Graphics, Color

public class ExtractImagesFromContentControls
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image that will be inserted into the document.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with a solid white color.
                g.Clear(Color.White);
            }
            // Save the bitmap so it can be used later.
            bitmap.Save(sampleImagePath);
        }

        // --------------------------------------------------------------
        // 2. Build a DOCX containing several content controls with images.
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to create a content control, insert a paragraph and an image.
        void InsertImageIntoContentControl(string tag, string title)
        {
            // Create a block‑level rich‑text content control.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
            sdt.Tag = tag;          // Identifier that will be used for the file name.
            sdt.Title = title;

            // Append the content control directly to the document body.
            doc.FirstSection.Body.AppendChild(sdt);

            // Content controls must contain a paragraph before an image can be added.
            Paragraph para = new Paragraph(doc);
            sdt.AppendChild(para);
            builder.MoveTo(para);

            // Insert the previously created sample image.
            builder.InsertImage(sampleImagePath);
        }

        // Create a few distinct content controls.
        InsertImageIntoContentControl("ControlA", "First control");
        InsertImageIntoContentControl("ControlB", "Second control");
        InsertImageIntoContentControl("ControlC", "Third control");

        // Save the document that now holds the images inside content controls.
        string docPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docPath);

        // --------------------------------------------------------------
        // 3. Load the document and extract images from each content control.
        // --------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        int totalExtracted = 0;

        // Get all StructuredDocumentTag nodes (content controls) in the document.
        NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Use the Tag property as the identifier for naming extracted files.
            string controlId = string.IsNullOrEmpty(sdt.Tag) ? $"Control_{sdt.Id}" : sdt.Tag;

            // Find all Shape nodes that are descendants of the current content control.
            NodeCollection shapeNodes = sdt.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    // Determine the appropriate file extension for the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{controlId}_{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(outputDir, imageFileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(imageFullPath);
                    imageIndex++;
                    totalExtracted++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the content controls.");

        // The program finishes automatically; all files are written to the Output folder.
    }
}
