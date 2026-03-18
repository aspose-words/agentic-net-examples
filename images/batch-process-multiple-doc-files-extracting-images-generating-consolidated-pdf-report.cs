using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace BatchImageExtractionAndPdfReport
{
    class Program
    {
        static void Main()
        {
            // Folder that contains the source DOC/DOCX files.
            string sourceFolder = @"C:\Docs\Input";
            // Folder where extracted images will be saved.
            string imagesFolder = @"C:\Docs\ExtractedImages";
            // Path for the final consolidated PDF report.
            string pdfReportPath = @"C:\Docs\ConsolidatedReport.pdf";

            // Ensure the images output folder exists.
            if (Directory.Exists(imagesFolder))
                Directory.Delete(imagesFolder, true);
            Directory.CreateDirectory(imagesFolder);

            // Create a new blank document that will become the PDF report.
            Document reportDoc = new Document();
            DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

            // Get all DOC and DOCX files from the source folder.
            string[] docFiles = Directory.GetFiles(sourceFolder, "*.doc*");

            foreach (string docPath in docFiles)
            {
                // Load each source document.
                Document srcDoc = new Document(docPath);

                // Collect all shapes (including images) from the source document.
                NodeCollection shapeNodes = srcDoc.GetChildNodes(NodeType.Shape, true);

                int imageIndex = 0;
                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    if (!shape.HasImage)
                        continue;

                    // Determine a unique file name for each extracted image.
                    string imageExtension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_{imageIndex}{imageExtension}";
                    string imageFullPath = Path.Combine(imagesFolder, imageFileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(imageFullPath);
                    imageIndex++;

                    // Insert the extracted image into the report document.
                    reportBuilder.InsertParagraph();
                    reportBuilder.InsertImage(imageFullPath);
                }

                // Add a page break after processing each source document (optional).
                reportBuilder.InsertBreak(BreakType.PageBreak);
            }

            // Save the consolidated report as a PDF file.
            reportDoc.Save(pdfReportPath, SaveFormat.Pdf);
        }
    }
}
