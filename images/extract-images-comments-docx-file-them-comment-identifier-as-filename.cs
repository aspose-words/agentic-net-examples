using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ExtractCommentImages <input.docx> <outputFolder>");
            return;
        }

        string inputFile = args[0];
        string outputFolder = args[1];

        try
        {
            ExtractCommentImages.Run(inputFile, outputFolder);
            Console.WriteLine("Images extracted successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

class ExtractCommentImages
{
    /// <summary>
    /// Extracts all images that are embedded inside comments of a DOCX file.
    /// Each image is saved to <paramref name="outputFolder"/> with a filename that
    /// contains the comment identifier and a sequential index.
    /// </summary>
    /// <param name="inputFile">Path to the source DOCX document.</param>
    /// <param name="outputFolder">Folder where extracted images will be saved.</param>
    public static void Run(string inputFile, string outputFolder)
    {
        // Load the document (lifecycle rule: use Document constructor)
        Document doc = new Document(inputFile);

        // Ensure the output directory exists (required for saving files)
        Directory.CreateDirectory(outputFolder);

        // Retrieve all comment nodes in the document
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through each comment
        foreach (Comment comment in comments)
        {
            // Find all Shape nodes that are descendants of the comment
            NodeCollection shapes = comment.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            // Process each shape that actually contains an image
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine the appropriate file extension for the image type
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    // Build a unique filename using the comment Id and image index
                    string fileName = $"Comment{comment.Id}_Img{imageIndex}{extension}";
                    string fullPath = Path.Combine(outputFolder, fileName);

                    // Save the image data to the file system (lifecycle rule: use ImageData.Save)
                    shape.ImageData.Save(fullPath);

                    imageIndex++;
                }
            }
        }
    }
}
