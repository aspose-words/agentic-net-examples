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
            // Path to the source DOCX file.
            const string inputDocPath = @"C:\Temp\InputDocument.docx";

            // Path to the output Excel‑compatible CSV file.
            const string outputCsvPath = @"C:\Temp\ImageMetadata.xlsx";

            // Load the Word document using the Document(string) constructor (load rule).
            Document doc = new Document(inputDocPath);

            // Collect all Shape nodes that contain images.
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            var imageInfos = new List<ImageInfo>();

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                if (!shape.HasImage) continue;

                // Gather metadata for each image.
                var info = new ImageInfo
                {
                    Index = imageIndex,
                    // Determine a file extension based on the image type.
                    Extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType),
                    // Size of the raw image data in bytes.
                    SizeInBytes = shape.ImageData.ImageBytes.Length,
                    // Width and height in points (as stored in the document).
                    WidthPoints = shape.Width,
                    HeightPoints = shape.Height,
                    // Original image type enum value.
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

            // Write the CSV content to a file with an .xlsx extension.
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
