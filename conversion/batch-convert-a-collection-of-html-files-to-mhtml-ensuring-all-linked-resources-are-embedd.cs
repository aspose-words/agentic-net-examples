using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = "InputHtml";
        string outputFolder = "OutputMhtml";

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample PNG image using Aspose.Drawing.
        string imagePath = Path.Combine(inputFolder, "image.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Fill the bitmap with a solid red color.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.Red);
            }

            // Save the bitmap as PNG using Aspose.Drawing.Imaging.ImageFormat.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Write a sample HTML file that references the image.
        string htmlFileName = "sample.html";
        string htmlFilePath = Path.Combine(inputFolder, htmlFileName);
        string htmlContent =
            "<html>" +
            "<body>" +
            "<h1>Sample Document</h1>" +
            "<p>Hello world!</p>" +
            "<img src=\"image.png\" alt=\"Sample Image\"/>" +
            "</body>" +
            "</html>";
        File.WriteAllText(htmlFilePath, htmlContent);

        // Batch convert each HTML file in the input folder to MHTML.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Configure save options for MHTML with embedded resources.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportCidUrlsForMhtmlResources = true, // Use CID URLs for resources.
                ExportFontResources = true,            // Export fonts if any.
                ExportImagesAsBase64 = false           // Keep images as separate parts.
            };

            // Determine the output MHTML file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mht";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as MHTML.
            doc.Save(outputPath, saveOptions);

            // Validate that the output file was created and is not empty.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected output MHTML file was not created: {outputPath}");

            FileInfo info = new FileInfo(outputPath);
            if (info.Length == 0)
                throw new InvalidOperationException($"Output MHTML file is empty: {outputPath}");
        }

        Console.WriteLine("Batch conversion completed successfully.");
    }
}
