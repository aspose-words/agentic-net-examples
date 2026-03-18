using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace InsertHighResPngExample
{
    class Program
    {
        static void Main()
        {
            // Path to the existing DOCX file.
            string docPath = @"C:\Docs\SourceDocument.docx";

            // Load the document using the Document(string) constructor.
            Document doc = new Document(docPath);

            // Create a DocumentBuilder for the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the target paragraph.
            // Parameters: paragraph index (zero‑based) and node index within that paragraph.
            // Adjust these values to point to the exact paragraph you need.
            int targetParagraphIndex = 2; // example: third paragraph in the body
            int targetNodeIndex = 0;      // start of the paragraph
            builder.MoveToParagraph(targetParagraphIndex, targetNodeIndex);

            // Path to the high‑resolution PNG image.
            string imagePath = @"C:\Images\HighResPicture.png";

            // Insert the image inline at 100 % scale.
            // This uses the InsertImage(string) overload as defined in the rules.
            Shape insertedImage = builder.InsertImage(imagePath);

            // Optional: adjust image size if required (width and height are in points).
            // insertedImage.Width = ConvertUtil.PixelToPoint(800);   // example width
            // insertedImage.Height = ConvertUtil.PixelToPoint(600);  // example height

            // Save the modified document to a new file.
            string outputPath = @"C:\Docs\ModifiedDocument.docx";
            doc.Save(outputPath);
        }
    }
}
