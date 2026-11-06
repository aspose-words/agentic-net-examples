using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output directories
        string inputDir = "InputDocs";
        string outputDir = "OutputHtml";

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files if none exist
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 2; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample document {i}");
                string inputPath = Path.Combine(inputDir, $"Sample{i}.docx");
                sampleDoc.Save(inputPath, SaveFormat.Docx);
                if (!File.Exists(inputPath))
                    throw new InvalidOperationException($"Failed to create sample DOCX: {inputPath}");
            }
        }

        // Batch convert each DOCX to HTML with round‑trip information enabled
        string[] docxFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            Document doc = new Document(docxPath);

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportRoundtripInformation = true
            };

            string htmlFileName = Path.GetFileNameWithoutExtension(docxPath) + ".html";
            string htmlPath = Path.Combine(outputDir, htmlFileName);

            doc.Save(htmlPath, saveOptions);

            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"HTML output was not created: {htmlPath}");
        }
    }
}
