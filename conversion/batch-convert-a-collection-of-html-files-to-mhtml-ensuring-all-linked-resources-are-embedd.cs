using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchHtmlToMhtmlConverter
{
    public static void Main()
    {
        // Define folders for input HTML files and output MHTML files.
        string baseDirectory = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDirectory, "InputHtml");
        string resourcesFolder = Path.Combine(inputFolder, "resources");
        string outputFolder = Path.Combine(baseDirectory, "OutputMhtml");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);

        Directory.CreateDirectory(resourcesFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample PNG image using Aspose.Drawing and save it to the resources folder.
        string imagePath = Path.Combine(resourcesFolder, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                graphics.FillEllipse(Brushes.DarkBlue, 10, 10, 80, 80);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create several HTML files that reference the image.
        for (int i = 1; i <= 3; i++)
        {
            string htmlFileName = $"sample{i}.html";
            string htmlFilePath = Path.Combine(inputFolder, htmlFileName);
            string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <title>Sample {i}</title>
</head>
<body>
    <h1>HTML Sample {i}</h1>
    <p>This is a sample HTML file number {i}.</p>
    <img src=""resources/sample.png"" alt=""Sample Image"" />
</body>
</html>";
            File.WriteAllText(htmlFilePath, htmlContent);
        }

        // Batch convert each HTML file to MHTML, embedding linked resources automatically.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Determine output MHTML file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mhtml";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save as MHTML. Resources are embedded by default.
            doc.Save(outputPath, SaveFormat.Mhtml);

            // Validate that the output file was created and is not empty.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");

            FileInfo info = new FileInfo(outputPath);
            if (info.Length == 0)
                throw new InvalidOperationException($"Output file '{outputPath}' is empty.");
        }

        // Indicate success.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
