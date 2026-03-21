using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Drawing;

class PdfImageExtractor
{
    static void Main()
    {
        // Folder containing the source PDF files (relative to the executable).
        string sourceFolder = Path.Combine(AppContext.BaseDirectory, "PdfSource");

        // Folder where extracted images will be saved (relative to the executable).
        string outputFolder = Path.Combine(AppContext.BaseDirectory, "PdfImages");

        // Ensure the directories exist.
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Get all PDF files in the source folder.
        string[] pdfFiles = Directory.GetFiles(sourceFolder, "*.pdf");

        if (pdfFiles.Length == 0)
        {
            Console.WriteLine($"No PDF files found in '{sourceFolder}'. Place PDFs there and rerun the program.");
            return;
        }

        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document using Aspose.Words.
            var loadOptions = new PdfLoadOptions();

            // Create a Document object from the PDF file.
            Document pdfDoc = new Document(pdfPath, loadOptions);

            // Determine a title for naming the extracted images.
            string docTitle = pdfDoc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(docTitle))
                docTitle = Path.GetFileNameWithoutExtension(pdfPath);

            // Collect all Shape nodes that may contain images.
            var shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);

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
                    Console.WriteLine($"Saved image: {imagePath}");

                    imageIndex++;
                }
            }

            if (imageIndex == 0)
                Console.WriteLine($"No images found in '{pdfPath}'.");
        }
    }
}
