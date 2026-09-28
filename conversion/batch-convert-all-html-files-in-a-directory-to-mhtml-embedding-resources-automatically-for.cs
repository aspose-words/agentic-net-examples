using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare input folder and sample HTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        Directory.CreateDirectory(inputFolder);

        // Sample HTML file 1.
        string htmlFile1 = Path.Combine(inputFolder, "sample1.html");
        File.WriteAllText(htmlFile1,
            "<html><body><h1>Sample 1</h1><p>This is the first sample HTML file.</p></body></html>");

        // Sample HTML file 2.
        string htmlFile2 = Path.Combine(inputFolder, "sample2.html");
        File.WriteAllText(htmlFile2,
            "<html><body><h1>Sample 2</h1><p>This is the second sample HTML file with an image.</p><img src='https://via.placeholder.com/150' /></body></html>");

        // Prepare output folder for MHTML files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputMhtml");
        Directory.CreateDirectory(outputFolder);

        // Batch convert each HTML file to MHTML.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Determine output file path with .mht extension.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".mht";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save as MHTML (resources are embedded automatically).
            doc.Save(outputPath, SaveFormat.Mhtml);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"MHTML file was not created: {outputPath}");
        }

        // Optional: indicate completion (no interactive wait).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
