using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

namespace AsposeWordsImageBatch
{
    public class Program
    {
        public static void Main()
        {
            // Set up folders.
            string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            string inputDir = Path.Combine(baseDir, "Input");
            string outputDir = Path.Combine(baseDir, "Output");
            string imagesDir = Path.Combine(outputDir, "Images");
            Directory.CreateDirectory(inputDir);
            Directory.CreateDirectory(outputDir);
            Directory.CreateDirectory(imagesDir);

            // Create sample images.
            string sampleImage1 = Path.Combine(baseDir, "sample1.png");
            string sampleImage2 = Path.Combine(baseDir, "sample2.png");
            CreateSampleImage(sampleImage1, 200, 150, Color.Blue);
            CreateSampleImage(sampleImage2, 300, 100, Color.Green);

            // Create sample DOCX files that contain the images.
            string doc1 = Path.Combine(inputDir, "Document1.docx");
            string doc2 = Path.Combine(inputDir, "Document2.docx");
            CreateSampleDocument(doc1, sampleImage1);
            CreateSampleDocument(doc2, sampleImage2);

            // Prepare CSV summary.
            StringBuilder csvBuilder = new StringBuilder();
            csvBuilder.AppendLine("DocumentName,ImageIndex,ImageFileName,WidthPixels,HeightPixels,ImageExtension");

            int totalExtractedImages = 0;

            // Process each DOCX file in the input folder.
            foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
            {
                Document doc = new Document(docPath);
                NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
                int imageIndex = 0;

                foreach (Shape shape in shapeNodes.OfType<Shape>())
                {
                    if (!shape.HasImage)
                        continue;

                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(imagesDir, imageFileName);

                    // Save the image.
                    shape.ImageData.Save(imageFullPath);

                    // Get image size in pixels.
                    ImageSize size = shape.ImageData.ImageSize;
                    int widthPx = size.WidthPixels;
                    int heightPx = size.HeightPixels;

                    // Append CSV line.
                    csvBuilder.AppendLine($"{Path.GetFileName(docPath)},{imageIndex},{imageFileName},{widthPx},{heightPx},{extension}");

                    imageIndex++;
                    totalExtractedImages++;
                }
            }

            // Validate that at least one image was extracted.
            if (totalExtractedImages == 0)
                throw new InvalidOperationException("No images were extracted from the DOCX files.");

            // Write CSV summary.
            string csvPath = Path.Combine(outputDir, "summary.csv");
            File.WriteAllText(csvPath, csvBuilder.ToString());

            // Clean up temporary sample images (optional).
            File.Delete(sampleImage1);
            File.Delete(sampleImage2);
        }

        // Creates a deterministic PNG image using Aspose.Drawing.
        private static void CreateSampleImage(string filePath, int width, int height, Color fillColor)
        {
            Bitmap bitmap = new Bitmap(width, height);
            Graphics graphics = Graphics.FromImage(bitmap);
            graphics.Clear(fillColor);
            bitmap.Save(filePath);
            graphics.Dispose();
            bitmap.Dispose();
        }

        // Creates a DOCX file that contains a single image.
        private static void CreateSampleDocument(string docPath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document generated for image: {Path.GetFileName(imagePath)}");
            builder.InsertImage(imagePath);
            doc.Save(docPath);
        }
    }
}
