using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputFolder = "InputHtml";
        string outputFolder = "OutputMhtml";

        // Ensure directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample image that will be referenced from the HTML files.
        string imagePath = Path.Combine(inputFolder, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Fill the bitmap with a solid color.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap as PNG.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create sample HTML files that reference the image.
        string htmlTemplate = @"<html>
  <body>
    <h1>Sample Document {0}</h1>
    <p>This is a sample HTML file.</p>
    <img src=""sample.png"" alt=""Sample Image"" />
  </body>
</html>";

        for (int i = 1; i <= 3; i++)
        {
            string htmlContent = string.Format(htmlTemplate, i);
            string htmlPath = Path.Combine(inputFolder, $"sample{i}.html");
            File.WriteAllText(htmlPath, htmlContent);
        }

        // Batch convert each HTML file to MHTML.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Determine the output MHTML path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mht";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save as MHTML (resources are embedded automatically).
            doc.Save(outputPath, SaveFormat.Mhtml);

            // Validate that the output file was created and is not empty.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");

            if (new FileInfo(outputPath).Length == 0)
                throw new InvalidOperationException($"Output file '{outputPath}' is empty.");
        }
    }
}
