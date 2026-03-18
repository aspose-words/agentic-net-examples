using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsImageToLatex
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Word document path
            string inputDocPath = @"C:\Input\Sample.docx";

            // Folder where extracted images will be saved
            string imagesFolder = @"C:\Output\Images";

            // Path of the generated LaTeX file
            string latexFilePath = @"C:\Output\DocumentImages.tex";

            // Ensure the images folder exists
            Directory.CreateDirectory(imagesFolder);

            // Load the Word document (lifecycle: load)
            Document doc = new Document(inputDocPath);

            // Collect all Shape nodes (including inline and floating images)
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            // StringBuilder to compose LaTeX content
            StringBuilder latexBuilder = new StringBuilder();

            // Optional LaTeX preamble (can be adjusted as needed)
            latexBuilder.AppendLine(@"\documentclass{article}");
            latexBuilder.AppendLine(@"\usepackage{graphicx}");
            latexBuilder.AppendLine(@"\begin{document}");
            latexBuilder.AppendLine();

            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on the image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build image file name and full path
                string imageFileName = $"image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesFolder, imageFileName);

                // Save the image data to the file system (lifecycle: save)
                shape.ImageData.Save(imageFullPath);

                // Add LaTeX code that references the saved image
                latexBuilder.AppendLine(@"\begin{figure}[h]");
                latexBuilder.AppendLine(@"\centering");
                latexBuilder.AppendLine($@"\includegraphics[width=\linewidth]{{{Path.Combine("Images", imageFileName)}}}");
                latexBuilder.AppendLine($@"\caption{{Image {imageIndex + 1}}}");
                latexBuilder.AppendLine(@"\end{figure}");
                latexBuilder.AppendLine();

                imageIndex++;
            }

            // Close LaTeX document
            latexBuilder.AppendLine(@"\end{document}");

            // Write the LaTeX content to the output .tex file
            File.WriteAllText(latexFilePath, latexBuilder.ToString());

            Console.WriteLine($"Extracted {imageIndex} image(s) to \"{imagesFolder}\".");
            Console.WriteLine($"LaTeX file generated at \"{latexFilePath}\".");
        }
    }
}
