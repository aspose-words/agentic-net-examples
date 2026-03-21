using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsImageExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Use folders relative to the executable directory so the example works out‑of‑the‑box
            string baseDir = AppContext.BaseDirectory;
            string docsFolder = Path.Combine(baseDir, "InputDocs");
            string imagesOutputFolder = Path.Combine(baseDir, "ExtractedImages");
            string csvReportPath = Path.Combine(baseDir, "ImageReport.csv");

            // Ensure the input folder exists; if not, create it and inform the user
            if (!Directory.Exists(docsFolder))
            {
                Directory.CreateDirectory(docsFolder);
                Console.WriteLine($"Created input folder at '{docsFolder}'. Place .docx files there and rerun the program.");
                return;
            }

            // Ensure output folder exists
            Directory.CreateDirectory(imagesOutputFolder);

            var csvRows = new List<string[]>();
            csvRows.Add(new[] { "DocumentPath", "ImageFileName", "ImageType", "WidthPoints", "HeightPoints", "SizeBytes" });

            // Process each .docx file in the input folder (non‑recursive)
            foreach (string docPath in Directory.GetFiles(docsFolder, "*.docx"))
            {
                ProcessDocument(docPath, imagesOutputFolder, csvRows);
            }

            // Write the CSV file
            WriteCsv(csvReportPath, csvRows);
            Console.WriteLine($"Processing complete. CSV report saved to '{csvReportPath}'.");
        }

        private static void ProcessDocument(string docPath, string imagesOutputFolder, List<string[]> csvRows)
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesOutputFolder, imageFileName);

                shape.ImageData.Save(imageFullPath);

                csvRows.Add(new[]
                {
                    docPath,
                    imageFileName,
                    shape.ImageData.ImageType.ToString(),
                    shape.Width.ToString(),
                    shape.Height.ToString(),
                    shape.ImageData.ImageBytes.Length.ToString()
                });

                imageIndex++;
            }
        }

        private static void WriteCsv(string csvPath, List<string[]> rows)
        {
            var sb = new StringBuilder();

            foreach (var row in rows)
            {
                string escaped = string.Join(",", row.Select(field =>
                {
                    string f = field.Replace("\"", "\"\"");
                    return $"\"{f}\"";
                }));
                sb.AppendLine(escaped);
            }

            File.WriteAllText(csvPath, sb.ToString(), Encoding.UTF8);
        }
    }
}
