using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Drawing; // <-- added namespace for Shape

class PdfImageExtractor
{
    static void Main()
    {
        // Folder containing the source PDF files.
        string sourceFolder = @"C:\PdfSource";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\PdfImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all PDF files in the source folder.
        string[] pdfFiles = Directory.GetFiles(sourceFolder, "*.pdf");

        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document using Aspose.Words.
            // The PdfLoadOptions class is used as the load options.
            PdfLoadOptions loadOptions = new PdfLoadOptions();

            // Create a Document object from the PDF file.
            Document pdfDoc = new Document(pdfPath, loadOptions);

            // Determine a title for naming the extracted images.
            // Prefer the document's Title property; fall back to the file name without extension.
            string docTitle = pdfDoc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(docTitle))
                docTitle = Path.GetFileNameWithoutExtension(pdfPath);

            // Collect all Shape nodes that may contain images.
            NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine the appropriate file extension for the image format.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    // Build a unique file name using the document title and an index.
                    string imageFileName = $"{docTitle}_image_{imageIndex}{extension}";
                    string imagePath = Path.Combine(outputFolder, imageFileName);

                    // Save the image data to the file system.
                    shape.ImageData.Save(imagePath);

                    imageIndex++;
                }
            }
        }
    }
}
