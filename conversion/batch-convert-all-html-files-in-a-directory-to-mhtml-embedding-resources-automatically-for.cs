using System;
using System.IO;
using Aspose.Words;

public class BatchHtmlToMhtmlConverter
{
    public static void Main()
    {
        // Define the input folder that will contain the sample HTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        Directory.CreateDirectory(inputFolder);

        // Create a couple of sample HTML files.
        string htmlFile1 = Path.Combine(inputFolder, "Sample1.html");
        string htmlFile2 = Path.Combine(inputFolder, "Sample2.html");

        File.WriteAllText(htmlFile1,
            "<html><body><h1>First Sample</h1><p>This is the first HTML file.</p></body></html>");
        File.WriteAllText(htmlFile2,
            "<html><body><h1>Second Sample</h1><p>This is the second HTML file.</p></body></html>");

        // Define the output folder for the generated MHTML files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputMhtml");
        Directory.CreateDirectory(outputFolder);

        // Process each HTML file in the input folder.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Determine the output MHTML file name (same base name, .mht extension).
            string outputFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".mht";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as MHTML. Resources (images, CSS, etc.) are embedded automatically.
            doc.Save(outputPath, SaveFormat.Mhtml);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"MHTML file was not created: {outputPath}");
        }

        // Optionally, indicate completion (no interactive prompts required).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
