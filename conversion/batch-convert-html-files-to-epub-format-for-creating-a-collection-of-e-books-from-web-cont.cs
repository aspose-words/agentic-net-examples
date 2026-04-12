using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output EPUB files.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "InputHtml");
        string outputFolder = Path.Combine(baseDir, "OutputEpub");

        // Ensure a clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        for (int i = 1; i <= 3; i++)
        {
            string htmlFile = Path.Combine(inputFolder, $"Sample{i}.html");
            string htmlContent = $@"<!DOCTYPE html>
<html>
<head><meta charset=""UTF-8""><title>Sample {i}</title></head>
<body>
<h1>Sample Document {i}</h1>
<p>This is a generated HTML file number {i}.</p>
</body>
</html>";
            File.WriteAllText(htmlFile, htmlContent);
        }

        // Batch convert each HTML file to EPUB.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Determine the output EPUB path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string epubPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.epub");

            // Save as EPUB using the native Aspose.Words API.
            doc.Save(epubPath, SaveFormat.Epub);

            // Validate that the EPUB file was created.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"Failed to create EPUB file: {epubPath}");
        }

        // Optional: list the generated EPUB files.
        Console.WriteLine("Batch conversion completed. Generated EPUB files:");
        foreach (string epub in Directory.GetFiles(outputFolder, "*.epub"))
        {
            Console.WriteLine(epub);
        }
    }
}
