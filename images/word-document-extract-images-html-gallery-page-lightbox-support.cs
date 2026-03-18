using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input Word document that contains images.
        const string inputDocPath = "Input.docx";

        // Output HTML file that will host the gallery.
        const string outputHtmlPath = "gallery.html";

        // Folder where extracted images will be stored.
        const string imagesFolder = "gallery_images";

        // -----------------------------------------------------------------
        // Prepare the images folder.
        // -----------------------------------------------------------------
        if (Directory.Exists(imagesFolder))
            Directory.Delete(imagesFolder, true);
        Directory.CreateDirectory(imagesFolder);

        // -----------------------------------------------------------------
        // Load the source Word document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document(inputDocPath);

        // -----------------------------------------------------------------
        // Extract all images from the document and save them to the folder.
        // Keep a list of the saved image file names for later HTML generation.
        // -----------------------------------------------------------------
        List<string> savedImageFiles = new List<string>();
        int imageIndex = 0;

        // Get all Shape nodes (they may contain images).
        NodeCollection shapes = sourceDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the proper file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build a unique file name.
                string fileName = $"img_{imageIndex}{extension}";
                string filePath = Path.Combine(imagesFolder, fileName);

                // Save the image data to the file system.
                shape.ImageData.Save(filePath);

                // Store the relative path for HTML generation.
                savedImageFiles.Add(fileName);
                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // Build the HTML gallery with Lightbox support.
        // -----------------------------------------------------------------
        var html = new System.Text.StringBuilder();
        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html>");
        html.AppendLine("<head>");
        html.AppendLine("<meta charset=\"UTF-8\">");
        html.AppendLine("<title>Image Gallery</title>");
        html.AppendLine("<link href=\"https://cdnjs.cloudflare.com/ajax/libs/lightbox2/2.11.3/css/lightbox.min.css\" rel=\"stylesheet\"/>");
        html.AppendLine("<style>");
        html.AppendLine(".gallery a { margin:5px; display:inline-block; }");
        html.AppendLine(".gallery img { width:150px; height:auto; border:1px solid #ccc; }");
        html.AppendLine("</style>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");
        html.AppendLine("<h1>Image Gallery</h1>");
        html.AppendLine("<div class=\"gallery\">");

        foreach (string fileName in savedImageFiles)
        {
            string imageUrl = $"{imagesFolder}/{fileName}";
            html.AppendLine($"<a href=\"{imageUrl}\" data-lightbox=\"gallery\"><img src=\"{imageUrl}\" alt=\"Image\"/></a>");
        }

        html.AppendLine("</div>");
        html.AppendLine("<script src=\"https://cdnjs.cloudflare.com/ajax/libs/lightbox2/2.11.3/js/lightbox.min.js\"></script>");
        html.AppendLine("</body>");
        html.AppendLine("</html>");

        File.WriteAllText(outputHtmlPath, html.ToString());
    }
}
