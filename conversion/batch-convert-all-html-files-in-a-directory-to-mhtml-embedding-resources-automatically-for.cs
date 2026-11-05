using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = "InputHtml";
        string outputFolder = "OutputMhtml";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files in the input folder.
        string sampleHtml1 = Path.Combine(inputFolder, "sample1.html");
        string sampleHtml2 = Path.Combine(inputFolder, "sample2.html");

        File.WriteAllText(sampleHtml1,
            "<html><head><title>Sample 1</title></head><body><h1>First Sample</h1><p>This is the first HTML file.</p></body></html>");
        File.WriteAllText(sampleHtml2,
            "<html><head><title>Sample 2</title></head><body><h1>Second Sample</h1><p>This is the second HTML file with an image.</p><img src='https://via.placeholder.com/150' alt='Placeholder'></body></html>");

        // Process each HTML file and convert it to MHTML.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Determine the output MHTML file path.
            string outputFileName = Path.ChangeExtension(Path.GetFileName(htmlPath), ".mhtml");
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save as MHTML, which embeds resources automatically.
            doc.Save(outputPath, SaveFormat.Mhtml);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected output MHTML was not created: {outputPath}");
        }

        // Optional: confirm that at least one file was processed.
        int outputCount = Directory.GetFiles(outputFolder, "*.mhtml").Length;
        if (outputCount == 0)
            throw new InvalidOperationException("No MHTML files were generated.");

        // Example completed successfully.
        Console.WriteLine($"Batch conversion completed. {outputCount} MHTML file(s) created in '{outputFolder}'.");
    }
}
