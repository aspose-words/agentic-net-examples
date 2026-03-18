using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;

class ExtractFootnoteImages
{
    static void Main()
    {
        // Load the source Word document.
        // (Assumes the document path is supplied; adjust as needed.)
        string inputPath = @"C:\Docs\InputDocument.docx";
        Document doc = new Document(inputPath);

        // Ensure that reference marks are up‑to‑date (required for correct numbering).
        doc.UpdateFields();
        doc.UpdateActualReferenceMarks();

        // Create a folder to store the extracted images.
        string outputFolder = @"C:\Docs\FootnoteImages";
        Directory.CreateDirectory(outputFolder);

        // Get all footnote nodes in the document.
        NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);

        foreach (Footnote footnote in footnoteNodes)
        {
            // Use the actual reference mark (e.g., "1", "2", …) as the footnote number.
            string footnoteNumber = footnote.ActualReferenceMark;

            // Find all image shapes inside the footnote.
            NodeCollection imageShapes = footnote.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Only process shapes that contain an image.
                if (shape.HasImage)
                {
                    // Build a unique file name: Footnote_<Number>_<Index>.<Extension>
                    string extension = shape.ImageData.ImageType.ToString().ToLower(); // e.g., "png", "jpeg"
                    string fileName = $"Footnote_{footnoteNumber}_{++imageIndex}.{extension}";
                    string filePath = Path.Combine(outputFolder, fileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(filePath);
                }
            }
        }

        // Optionally, save the (unchanged) document to a new file.
        // This demonstrates the required save lifecycle step.
        string outputDocPath = @"C:\Docs\ProcessedDocument.docx";
        doc.Save(outputDocPath);
    }
}
