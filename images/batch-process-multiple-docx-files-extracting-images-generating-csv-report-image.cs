using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace AsposeWordsImageExtractor
{
    class Program
    {
        // Entry point
        static void Main(string[] args)
        {
            // Folder that contains the DOCX files to process
            string docsFolder = @"C:\InputDocs";

            // Folder where extracted images will be saved
            string imagesOutputFolder = @"C:\ExtractedImages";

            // Path of the CSV report to generate
            string csvReportPath = @"C:\ImageReport.csv";

            // Ensure output folder exists
            Directory.CreateDirectory(imagesOutputFolder);

            // Prepare a list to hold CSV rows (each row is an array of string values)
            var csvRows = new List<string[]>();
            // Add header row
            csvRows.Add(new[] { "DocumentPath", "ImageFileName", "ImageType", "WidthPoints", "HeightPoints", "SizeBytes" });

            // Process each .docx file in the input folder (non‑recursive)
            foreach (string docPath in Directory.GetFiles(docsFolder, "*.docx"))
            {
                ProcessDocument(docPath, imagesOutputFolder, csvRows);
            }

            // Write the CSV file
            WriteCsv(csvReportPath, csvRows);
        }

        // Loads a document, extracts all images, saves them, and records metadata.
        private static void ProcessDocument(string docPath, string imagesOutputFolder, List<string[]> csvRows)
        {
            // Load the document using the Document(string) constructor (load rule)
            Document doc = new Document(docPath);

            // Get all Shape nodes (including images) from the document
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue; // Skip shapes that are not images

                // Determine a file extension based on the image type (utility rule)
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build a unique file name for the extracted image
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesOutputFolder, imageFileName);

                // Save the image to the file system (uses ImageData.Save method)
                shape.ImageData.Save(imageFullPath);

                // Gather metadata
                string imageType = shape.ImageData.ImageType.ToString();
                // Width and Height are stored in points (1 point = 1/72 inch)
                string width = shape.Width.ToString();
                string height = shape.Height.ToString();
                // Size in bytes of the raw image data
                string sizeBytes = shape.ImageData.ImageBytes.Length.ToString();

                // Add a row to the CSV data
                csvRows.Add(new[]
                {
                    docPath,
                    imageFileName,
                    imageType,
                    width,
                    height,
                    sizeBytes
                });

                imageIndex++;
            }
        }

        // Writes the collected CSV rows to a file using UTF‑8 encoding.
        private static void WriteCsv(string csvPath, List<string[]> rows)
        {
            var sb = new StringBuilder();

            foreach (var row in rows)
            {
                // Escape any double quotes in fields and wrap each field in quotes
                string escaped = string.Join(",", row.Select(field =>
                {
                    string f = field.Replace("\"", "\"\"");
                    return $"\"{f}\"";
                }));
                sb.AppendLine(escaped);
            }

            // Save the CSV file (standard file I/O, not a Document save operation)
            File.WriteAllText(csvPath, sb.ToString(), Encoding.UTF8);
        }
    }
}
