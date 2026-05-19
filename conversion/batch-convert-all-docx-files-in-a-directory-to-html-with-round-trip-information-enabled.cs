using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputHtml");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Seed the input folder with a few sample DOCX files.
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample document {i}");
            string docxPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            sampleDoc.Save(docxPath, SaveFormat.Docx);
        }

        // Convert each DOCX file to HTML with round‑trip information enabled.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docxPath);
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportRoundtripInformation = true
            };

            string htmlPath = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(docxPath) + ".html");
            doc.Save(htmlPath, htmlOptions);

            // Verify that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to create HTML file: {htmlPath}");
        }
    }
}
