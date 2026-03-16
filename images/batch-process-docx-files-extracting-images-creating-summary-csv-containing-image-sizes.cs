using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsImageExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing DOCX files.
            string inputFolder = @"C:\Docs\Input";
            // Output folder where extracted images will be saved.
            string imagesOutputFolder = @"C:\Docs\ExtractedImages";
            // Path to the summary CSV file.
            string csvSummaryPath = @"C:\Docs\ImageSummary.csv";

            // Ensure the images output folder exists.
            Directory.CreateDirectory(imagesOutputFolder);

            // Prepare CSV writer.
            using (var csvWriter = new StreamWriter(csvSummaryPath, false))
            {
                // Write CSV header.
                csvWriter.WriteLine("Document,ImageFile,WidthPoints,HeightPoints,SizeBytes");

                // Process each DOCX file in the input folder.
                foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
                {
                    // Load the document using the provided constructor.
                    Document doc = new Document(docxPath);

                    // Retrieve all Shape nodes (including those inside headers/footers).
                    NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

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

                        // Determine image size in bytes (using the saved file).
                        long sizeBytes = new FileInfo(imageFullPath).Length;

                        // Write a line to the CSV summary.
                        csvWriter.WriteLine($"{Path.GetFileName(docxPath)},{imageFileName},{shape.Width},{shape.Height},{sizeBytes}");

                        imageIndex++;
                    }
                }
            }

            Console.WriteLine("Image extraction and CSV summary completed.");
        }
    }
}
