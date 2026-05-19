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

        // Create a simple image that will be referenced from the HTML files.
        string imagePath = Path.Combine(inputFolder, "sample.png");
        CreateSampleImage(imagePath);

        // Create a few sample HTML files that reference the image.
        for (int i = 1; i <= 2; i++)
        {
            string htmlContent = $@"<html>
    <body>
        <h1>Sample Document {i}</h1>
        <p>This is a test HTML file.</p>
        <img src='sample.png' alt='Sample Image' />
    </body>
</html>";
            string htmlFilePath = Path.Combine(inputFolder, $"sample{i}.html");
            File.WriteAllText(htmlFilePath, htmlContent);
        }

        // Batch convert each HTML file to MHTML with embedded resources.
        foreach (string htmlFile in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Prepare the output MHTML file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mht";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Configure save options for MHTML. Enable CID URLs to ensure resources are embedded.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportCidUrlsForMhtmlResources = true
            };

            // Save the document as MHTML.
            doc.Save(outputPath, saveOptions);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create MHTML file: {outputPath}");
        }

        // Optionally, indicate successful conversion (no interactive output required).
        // Console.WriteLine("Batch conversion completed.");
    }

    // Creates a simple 100x100 PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string path)
    {
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                graphics.DrawEllipse(new Pen(Color.DarkBlue, 3), 10, 10, 80, 80);
            }

            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
