using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = "InputHtml";
        string outputDir = "OutputMhtml";

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Seed the input folder with sample HTML files.
        CreateSampleHtml(Path.Combine(inputDir, "sample1.html"),
            "<html><body><h1>Sample 1</h1><p>Hello World!</p></body></html>");

        CreateSampleHtml(Path.Combine(inputDir, "sample2.html"),
            "<html><body><h1>Sample 2</h1><img src=\"https://via.placeholder.com/150\" alt=\"Placeholder\"/></body></html>");

        // Convert each HTML file in the input folder to MHTML.
        foreach (string htmlFilePath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlFilePath);

            // Build the output MHTML file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlFilePath);
            string mhtmlFilePath = Path.Combine(outputDir, fileNameWithoutExt + ".mht");

            // Save as MHTML; resources (images, fonts, CSS) are embedded automatically.
            doc.Save(mhtmlFilePath, SaveFormat.Mhtml);

            // Verify that the output file was created.
            if (!File.Exists(mhtmlFilePath))
                throw new InvalidOperationException($"Failed to create MHTML file: {mhtmlFilePath}");
        }
    }

    // Helper method to write HTML content to a file.
    private static void CreateSampleHtml(string filePath, string htmlContent)
    {
        File.WriteAllText(filePath, htmlContent);
    }
}
