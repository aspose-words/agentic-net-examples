using System;
using System.IO;
using System.Text;
using System.Linq;
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
            // Create a simple Word document in memory (no external file needed)
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document without images.");

            // Folder where extracted images will be saved (relative to the current directory)
            string imagesFolder = Path.Combine(Environment.CurrentDirectory, "Output", "Images");

            // Path of the generated LaTeX file (relative to the current directory)
            string latexFilePath = Path.Combine(Environment.CurrentDirectory, "Output", "DocumentImages.tex");

            // Ensure the images folder exists
            Directory.CreateDirectory(imagesFolder);
            // Ensure the output folder exists
            Directory.CreateDirectory(Path.GetDirectoryName(latexFilePath)!);

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

                // Save the image data to the file system
                shape.ImageData.Save(imageFullPath);

                // Add LaTeX code that references the saved image
                latexBuilder.AppendLine(@"\begin{figure}[h]");
                latexBuilder.AppendLine(@"\centering");
                latexBuilder.AppendLine($@"\includegraphics[width=\linewidth]{{Images/{imageFileName}}}");
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
