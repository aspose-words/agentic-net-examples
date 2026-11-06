using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchHtmlToMhtmlConverter
{
    public static void Main()
    {
        // Define input and output folders
        string inputFolder = "InputHtml";
        string outputFolder = "OutputMhtml";

        // Ensure clean start
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample image that will be referenced from HTML files
        string imagePath = Path.Combine(inputFolder, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a few sample HTML files referencing the image
        for (int i = 1; i <= 3; i++)
        {
            string htmlFileName = $"page{i}.html";
            string htmlPath = Path.Combine(inputFolder, htmlFileName);
            string htmlContent = $@"
<!DOCTYPE html>
<html>
<head><title>Sample Page {i}</title></head>
<body>
<h1>Sample Page {i}</h1>
<p>This is a sample HTML file.</p>
<img src=""sample.png"" alt=""Sample Image"" />
</body>
</html>";
            File.WriteAllText(htmlPath, htmlContent);
        }

        // Batch convert each HTML file to MHTML
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load HTML document
            Document doc = new Document(htmlFile);

            // Determine output MHTML path
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mhtml";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save as MHTML (resources are embedded automatically)
            doc.Save(outputPath, SaveFormat.Mhtml);

            // Validate output
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"MHTML file was not created: {outputPath}");

            FileInfo info = new FileInfo(outputPath);
            if (info.Length == 0)
                throw new InvalidOperationException($"MHTML file is empty: {outputPath}");
        }

        // Optional: indicate success (no interactive output required)
        // The program will exit normally if all conversions succeed.
    }
}
