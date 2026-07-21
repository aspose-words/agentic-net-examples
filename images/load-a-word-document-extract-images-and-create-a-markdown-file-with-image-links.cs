using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

namespace AsposeWordsImageExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a deterministic sample image (input.png).
            // -----------------------------------------------------------------
            string sampleImagePath = Path.Combine(outputDir, "input.png");
            CreateSampleImage(sampleImagePath);

            // -----------------------------------------------------------------
            // 2. Create a Word document and insert the sample image.
            // -----------------------------------------------------------------
            string docPath = Path.Combine(outputDir, "sample.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);

            // -----------------------------------------------------------------
            // 3. Load the document and extract all images to separate files.
            // -----------------------------------------------------------------
            Document doc = new Document(docPath);
            List<string> extractedImageFiles = ExtractImages(doc, outputDir);

            // Validate that at least one image was extracted.
            if (extractedImageFiles.Count == 0)
                throw new InvalidOperationException("No images were extracted from the document.");

            // -----------------------------------------------------------------
            // 4. Generate a Markdown file that references the extracted images.
            // -----------------------------------------------------------------
            string markdownPath = Path.Combine(outputDir, "Images.md");
            GenerateMarkdownFile(markdownPath, extractedImageFiles);
        }

        // Creates a simple 100x100 PNG image with a solid background.
        private static void CreateSampleImage(string filePath)
        {
            const int width = 100;
            const int height = 100;

            // Use Aspose.Drawing to create a bitmap and fill it with a color.
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            graphics.Clear(Aspose.Drawing.Color.LightBlue);

            // Save the bitmap to the specified file.
            bitmap.Save(filePath);
            graphics.Dispose();
            bitmap.Dispose();
        }

        // Creates a blank document and inserts the image located at imagePath.
        private static void CreateDocumentWithImage(string docPath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(imagePath);
            doc.Save(docPath);
        }

        // Extracts all images from the provided document into the output directory.
        // Returns a list of file names (relative to the output directory) of the saved images.
        private static List<string> ExtractImages(Document doc, string outputDir)
        {
            List<string> savedFiles = new List<string>();
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string fileName = $"image{imageIndex}{extension}";
                    string fullPath = Path.Combine(outputDir, fileName);
                    shape.ImageData.Save(fullPath);
                    savedFiles.Add(fileName);
                    imageIndex++;
                }
            }

            return savedFiles;
        }

        // Writes a Markdown file where each line contains an image link to the provided files.
        private static void GenerateMarkdownFile(string markdownPath, List<string> imageFiles)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string img in imageFiles)
            {
                sb.AppendLine($"![]({img})");
            }

            File.WriteAllText(markdownPath, sb.ToString(), Encoding.UTF8);
        }
    }
}
