using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchHtmlToMhtmlConverter
{
    public static void Main()
    {
        // Create a temporary folder for sample HTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "HtmlSamples");
        Directory.CreateDirectory(inputFolder);

        // Create a sample image that will be referenced from the HTML.
        string imagePath = Path.Combine(inputFolder, "sampleImage.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a sample HTML file that references the image.
        string htmlContent = @"<html>
    <head><title>Sample HTML</title></head>
    <body>
        <h1>Sample Document</h1>
        <p>This is a sample HTML file with an embedded image.</p>
        <img src=""sampleImage.png"" alt=""Sample Image"" />
    </body>
</html>";
        string htmlFilePath = Path.Combine(inputFolder, "sample1.html");
        File.WriteAllText(htmlFilePath, htmlContent);

        // Create another simple HTML file without external resources.
        string htmlContent2 = @"<html>
    <head><title>Second Sample</title></head>
    <body>
        <h2>Second Document</h2>
        <p>Just some text.</p>
    </body>
</html>";
        string htmlFilePath2 = Path.Combine(inputFolder, "sample2.html");
        File.WriteAllText(htmlFilePath2, htmlContent2);

        // Batch convert all HTML files in the folder to MHTML.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Determine the output MHTML file path.
            string mhtmlFile = Path.ChangeExtension(htmlFile, ".mht");

            // Save as MHTML. Resources (images, CSS, etc.) are embedded automatically.
            doc.Save(mhtmlFile, SaveFormat.Mhtml);

            // Validate that the output file was created and is not empty.
            if (!File.Exists(mhtmlFile))
                throw new InvalidOperationException($"Failed to create MHTML file: {mhtmlFile}");

            FileInfo info = new FileInfo(mhtmlFile);
            if (info.Length == 0)
                throw new InvalidOperationException($"MHTML file is empty: {mhtmlFile}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
