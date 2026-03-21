using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsImageExtractor
{
    class Program
    {
        static void Main()
        {
            // Use a temporary folder for all files so the example works on any machine.
            string tempFolder = Path.GetTempPath();

            // Path to the source DOCX file.
            string inputDocPath = Path.Combine(tempFolder, "InputDocument.docx");

            // Path to the output CSV file (Excel can open .csv files).
            string outputCsvPath = Path.Combine(tempFolder, "ImageMetadata.csv");

            // Ensure a document exists. If the file is missing, create a simple one.
            Document doc;
            if (File.Exists(inputDocPath))
            {
                doc = new Document(inputDocPath);
            }
            else
            {
                doc = new Document();
                // Add a paragraph so the document is not empty.
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln("Sample document created by AsposeWordsImageExtractor.");
                // Save the placeholder document for future runs.
                doc.Save(inputDocPath);
            }

            // Collect all Shape nodes that contain images.
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            var imageInfos = new List<ImageInfo>();

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                if (!shape.HasImage) continue;

                var info = new ImageInfo
                {
                    Index = imageIndex,
                    Extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType),
                    SizeInBytes = shape.ImageData.ImageBytes.Length,
                    WidthPoints = shape.Width,
                    HeightPoints = shape.Height,
                    ImageType = shape.ImageData.ImageType.ToString()
                };
                imageInfos.Add(info);
                imageIndex++;
            }

            // Build a CSV string that Excel can open.
            var sb = new StringBuilder();
            sb.AppendLine("Index,Extension,ImageType,SizeInBytes,WidthPoints,HeightPoints");
            foreach (var info in imageInfos)
            {
                sb.AppendLine($"{info.Index},{info.Extension},{info.ImageType},{info.SizeInBytes},{info.WidthPoints},{info.HeightPoints}");
            }

            // Write the CSV content.
            File.WriteAllText(outputCsvPath, sb.ToString(), Encoding.UTF8);

            Console.WriteLine($"Extracted {imageInfos.Count} image(s) and wrote metadata to '{outputCsvPath}'.");
        }

        // Simple DTO to hold image metadata.
        private class ImageInfo
        {
            public int Index { get; set; }
            public string Extension { get; set; }
            public string ImageType { get; set; }
            public int SizeInBytes { get; set; }
            public double WidthPoints { get; set; }
            public double HeightPoints { get; set; }
        }
    }
}
