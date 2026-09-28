using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDirectory = "InputDocs";
        string outputDirectory = "OutputHtml";

        // Ensure directories exist.
        Directory.CreateDirectory(inputDirectory);
        Directory.CreateDirectory(outputDirectory);

        // Create sample DOCX files in the input directory.
        for (int i = 1; i <= 2; i++)
        {
            string docxPath = Path.Combine(inputDirectory, $"Sample{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document {i}.");
            sampleDoc.Save(docxPath, SaveFormat.Docx);
        }

        // Batch convert each DOCX file to HTML with round‑trip information enabled.
        string[] docxFiles = Directory.GetFiles(inputDirectory, "*.docx");
        foreach (string docxFilePath in docxFiles)
        {
            Document doc = new Document(docxFilePath);

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(docxFilePath);
            string htmlFilePath = Path.Combine(outputDirectory, fileNameWithoutExtension + ".html");

            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.ExportRoundtripInformation = true;

            doc.Save(htmlFilePath, htmlOptions);

            if (!File.Exists(htmlFilePath))
                throw new InvalidOperationException($"Expected HTML file was not created: {htmlFilePath}");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
