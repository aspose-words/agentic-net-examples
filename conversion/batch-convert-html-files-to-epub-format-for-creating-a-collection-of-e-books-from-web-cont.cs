using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputHtml");
        string outputDir = Path.Combine(baseDir, "OutputEpub");

        // Ensure a clean environment.
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few deterministic HTML files to act as source e‑books.
        CreateSampleHtml(Path.Combine(inputDir, "Sample1.html"),
            "<html><body><h1>First Book</h1><p>Content of the first e‑book.</p></body></html>");
        CreateSampleHtml(Path.Combine(inputDir, "Sample2.html"),
            "<html><body><h1>Second Book</h1><p>More content here.</p></body></html>");

        // Iterate over each HTML file in the input folder and convert it to EPUB.
        foreach (string htmlPath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Configure EPUB save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Epub;
            saveOptions.Encoding = Encoding.UTF8;
            // Optional: split the EPUB into parts at heading paragraphs.
            // saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

            // Build the output EPUB file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string epubPath = Path.Combine(outputDir, fileNameWithoutExt + ".epub");

            // Perform the conversion.
            doc.Save(epubPath, saveOptions);

            // Validate that the EPUB file was created.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to write an HTML string to a file using UTF‑8 encoding.
    private static void CreateSampleHtml(string path, string htmlContent)
    {
        File.WriteAllText(path, htmlContent, Encoding.UTF8);
    }
}
