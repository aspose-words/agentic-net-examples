using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsImageExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Base directory for the demo (located next to the executable).
            string baseDir = AppContext.BaseDirectory;

            // Input folder containing DOCX files.
            string inputFolder = Path.Combine(baseDir, "Input");
            // Output folder where extracted images will be saved.
            string imagesOutputFolder = Path.Combine(baseDir, "ExtractedImages");
            // Path to the summary CSV file.
            string csvSummaryPath = Path.Combine(baseDir, "ImageSummary.csv");

            // Ensure the required folders exist.
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(imagesOutputFolder);

            // If there are no DOCX files, inform the user and exit gracefully.
            var docxFiles = Directory.GetFiles(inputFolder, "*.docx");
            if (docxFiles.Length == 0)
            {
                Console.WriteLine($"No DOCX files found in '{inputFolder}'. Place files there and rerun the program.");
                return;
            }

            // Prepare CSV writer.
            using (var csvWriter = new StreamWriter(csvSummaryPath, false))
            {
                // Write CSV header.
                csvWriter.WriteLine("Document,ImageFile,WidthPoints,HeightPoints,SizeBytes");

                // Process each DOCX file in the input folder.
                foreach (string docxPath in docxFiles)
                {
                    // Load the document.
                    Document doc = new Document(docxPath);

                    // Retrieve all Shape nodes (including those inside headers/footers).
                    var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

                    int imageIndex = 0;
                    foreach (Shape shape in shapeNodes.OfType<Shape>())
                    {
                        if (!shape.HasImage)
                            continue; // Skip shapes without images.

                        // Build a unique image file name.
                        string docNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
                        string imageExtension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                        string imageFileName = $"{docNameWithoutExt}_img{imageIndex}{imageExtension}";
                        string imageFullPath = Path.Combine(imagesOutputFolder, imageFileName);

                        // Save the image to the file system.
                        shape.ImageData.Save(imageFullPath);

                        // Determine image size in bytes.
                        long sizeBytes = new FileInfo(imageFullPath).Length;

                        // Write a line to the CSV summary.
                        csvWriter.WriteLine($"{Path.GetFileName(docxPath)},{imageFileName},{shape.Width},{shape.Height},{sizeBytes}");

                        imageIndex++;
                    }
                }
            }

            Console.WriteLine("Image extraction and CSV summary completed.");
            Console.WriteLine($"Images saved to: {imagesOutputFolder}");
            Console.WriteLine($"CSV summary saved to: {csvSummaryPath}");
        }
    }
}
