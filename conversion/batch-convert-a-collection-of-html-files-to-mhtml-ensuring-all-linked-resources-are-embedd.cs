using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchHtmlToMhtmlConverter
{
    public static void Main()
    {
        // Define folders for input HTML files and output MHTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputMhtml");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample resources (a tiny PNG image) that will be referenced from the HTML files.
        string imagePath = Path.Combine(inputFolder, "sampleImage.png");
        CreateSamplePng(imagePath);

        // Create a few sample HTML files that reference the image.
        for (int i = 1; i <= 2; i++)
        {
            string htmlFileName = $"sample{i}.html";
            string htmlFilePath = Path.Combine(inputFolder, htmlFileName);
            string htmlContent = $@"<html>
    <body>
        <h1>Sample Document {i}</h1>
        <p>This is a test HTML file.</p>
        <img src=""sampleImage.png"" alt=""Sample Image"" />
    </body>
</html>";
            File.WriteAllText(htmlFilePath, htmlContent);
        }

        // Batch convert each HTML file in the input folder to MHTML.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Determine the output MHTML file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mht";
            string outputFilePath = Path.Combine(outputFolder, outputFileName);

            // Save the document as MHTML. Resources (images, CSS, etc.) are embedded automatically.
            doc.Save(outputFilePath, SaveFormat.Mhtml);

            // Verify that the output file was created.
            if (!File.Exists(outputFilePath))
            {
                throw new InvalidOperationException($"Failed to create MHTML file: {outputFilePath}");
            }
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Creates a minimal 1x1 pixel PNG file without using System.Drawing.
    private static void CreateSamplePng(string filePath)
    {
        // PNG file header and a single transparent pixel.
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A, // PNG signature
            0x00,0x00,0x00,0x0D, // IHDR chunk length
            0x49,0x48,0x44,0x52, // "IHDR"
            0x00,0x00,0x00,0x01, // width: 1
            0x00,0x00,0x00,0x01, // height: 1
            0x08, // bit depth
            0x06, // color type: RGBA
            0x00, // compression method
            0x00, // filter method
            0x00, // interlace method
            0x1F,0x15,0xC4,0x89, // CRC
            0x00,0x00,0x00,0x0A, // IDAT chunk length
            0x49,0x44,0x41,0x54, // "IDAT"
            0x78,0x9C,0x63,0x60,0x00,0x00,0x00,0x02,0x00,0x01, // compressed data
            0x05,0x5C,0x0A,0x2D, // CRC
            0x00,0x00,0x00,0x00, // IEND chunk length
            0x49,0x45,0x4E,0x44, // "IEND"
            0xAE,0x42,0x60,0x82  // CRC
        };
        File.WriteAllBytes(filePath, pngBytes);
    }
}
