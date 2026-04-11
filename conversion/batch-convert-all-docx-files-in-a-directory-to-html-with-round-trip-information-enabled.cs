using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchDocxToHtml
{
    public static void Main()
    {
        // Define folders for input DOCX files and output HTML files.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputHtml");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files if the input folder is empty.
        CreateSampleDocxIfMissing(inputDir, "Sample1.docx", "This is the first sample document.");
        CreateSampleDocxIfMissing(inputDir, "Sample2.docx", "This is the second sample document.");

        // Process each DOCX file in the input directory.
        string[] docxFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Configure HTML save options with round‑trip information enabled.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportRoundtripInformation = true,
                // Save any images to the same output folder.
                ImagesFolder = outputDir
            };

            // Determine the output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(docxPath) + ".html";
            string htmlPath = Path.Combine(outputDir, htmlFileName);

            // Save the document as HTML.
            doc.Save(htmlPath, htmlOptions);

            // Verify that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to create HTML file: {htmlPath}");
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple DOCX file if it does not already exist.
    private static void CreateSampleDocxIfMissing(string folder, string fileName, string content)
    {
        string fullPath = Path.Combine(folder, fileName);
        if (File.Exists(fullPath))
            return;

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(fullPath, SaveFormat.Docx);
    }
}
