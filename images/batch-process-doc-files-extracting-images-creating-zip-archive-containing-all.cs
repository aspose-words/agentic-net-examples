using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Folder that contains the source DOC/DOCX files.
        string sourceFolder = @"C:\Docs";

        // Temporary folder where extracted images will be saved.
        string tempImageFolder = Path.Combine(Path.GetTempPath(), "AsposeExtractedImages");
        if (Directory.Exists(tempImageFolder))
            Directory.Delete(tempImageFolder, true);
        Directory.CreateDirectory(tempImageFolder);

        // Process each Word document in the source folder.
        foreach (string docPath in Directory.GetFiles(sourceFolder, "*.doc*"))
        {
            // Load the document using the provided constructor rule.
            Document doc = new Document(docPath);

            // Retrieve all Shape nodes (including images) from the document.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    // Determine the proper file extension for the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    // Build a unique file name that includes the source document name.
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(tempImageFolder, imageFileName);

                    // Save the image to the temporary folder.
                    shape.ImageData.Save(imageFullPath);
                    imageIndex++;
                }
            }
        }

        // Path for the resulting ZIP archive.
        string zipPath = Path.Combine(sourceFolder, "ExtractedImages.zip");

        // Remove any existing archive with the same name.
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        // Create a ZIP archive containing all extracted images.
        ZipFile.CreateFromDirectory(tempImageFolder, zipPath);
    }
}
