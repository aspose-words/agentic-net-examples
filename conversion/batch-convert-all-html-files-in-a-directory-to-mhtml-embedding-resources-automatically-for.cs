using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchHtmlToMhtmlConverter
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputMhtml");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Seed sample HTML files.
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample1.html"), "<html><body><h1>Sample 1</h1><p>Hello World!</p></body></html>");
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample2.html"), "<html><body><h1>Sample 2</h1><p>Another document.</p></body></html>");

        // Process each HTML file in the input directory.
        foreach (string htmlPath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Determine output MHTML path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string mhtmlPath = Path.Combine(outputDir, fileNameWithoutExt + ".mhtml");

            // Save as MHTML (resources are embedded automatically).
            doc.Save(mhtmlPath, SaveFormat.Mhtml);

            // Validate that the output file was created.
            if (!File.Exists(mhtmlPath))
                throw new InvalidOperationException($"MHTML file was not created: {mhtmlPath}");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    private static void CreateSampleHtmlFile(string path, string content)
    {
        // Write deterministic HTML content to the specified file.
        File.WriteAllText(path, content);
    }
}
