using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Properties;

class ExtractCustomPropertyImages
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Docs\ExtractedImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the Word document (lifecycle rule: use Document constructor).
        Document doc = new Document(inputPath);

        // Iterate through all custom document properties.
        foreach (DocumentProperty prop in doc.CustomDocumentProperties)
        {
            // Check if the property stores a byte array (potential image data).
            if (prop.Type == PropertyType.ByteArray)
            {
                // Retrieve the raw bytes of the property.
                byte[] imageBytes = prop.ToByteArray();

                // Determine a suitable file extension based on the image format.
                string extension = GetImageExtension(imageBytes);

                // Build a safe file name using the property name.
                string safeName = MakeFileSystemSafeName(prop.Name);
                string filePath = Path.Combine(outputFolder, safeName + extension);

                // Save the image bytes to disk (lifecycle rule: use File.WriteAllBytes).
                File.WriteAllBytes(filePath, imageBytes);

                Console.WriteLine($"Saved image from property \"{prop.Name}\" to \"{filePath}\"");
            }
        }
    }

    // Returns a file extension (including the dot) that matches the image data.
    // Uses simple magic‑number detection to avoid System.Drawing dependencies.
    private static string GetImageExtension(byte[] data)
    {
        if (data == null || data.Length < 4)
            return ".bin";

        // JPEG: FF D8 FF
        if (data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF)
            return ".jpg";
        // PNG: 89 50 4E 47
        if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
            return ".png";
        // GIF: 47 49 46 38
        if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x38)
            return ".gif";
        // BMP: 42 4D
        if (data[0] == 0x42 && data[1] == 0x4D)
            return ".bmp";
        // TIFF (little endian): 49 49 2A 00
        if (data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00)
            return ".tiff";
        // TIFF (big endian): 4D 4D 00 2A
        if (data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A)
            return ".tiff";

        // Unknown format – fall back to generic binary extension.
        return ".bin";
    }

    // Replaces characters that are invalid in file names.
    private static string MakeFileSystemSafeName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
