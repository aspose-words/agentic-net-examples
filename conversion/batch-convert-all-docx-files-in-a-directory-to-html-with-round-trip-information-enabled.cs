using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Prepare input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputHtml");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files if they do not already exist.
        CreateSampleDocx(Path.Combine(inputDir, "Sample1.docx"), "First sample document.", "This is the first sample.");
        CreateSampleDocx(Path.Combine(inputDir, "Sample2.docx"), "Second sample document.", "This is the second sample.");

        // Get all DOCX files in the input directory.
        string[] docxFiles = Directory.GetFiles(inputDir, "*.docx", SearchOption.TopDirectoryOnly);
        if (docxFiles.Length == 0)
        {
            Console.WriteLine("No DOCX files found to convert.");
            return;
        }

        // Convert each DOCX to HTML with round‑trip information enabled (if supported).
        foreach (string docxPath in docxFiles)
        {
            Document doc = new Document(docxPath);

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Enable round‑trip information when the property exists in the current Aspose.Words version.
            PropertyInfo? roundTripProp = typeof(HtmlSaveOptions).GetProperty("ExportRoundTripInformation", BindingFlags.Public | BindingFlags.Instance);
            if (roundTripProp != null && roundTripProp.CanWrite)
            {
                roundTripProp.SetValue(saveOptions, true);
            }

            string outputFileName = Path.GetFileNameWithoutExtension(docxPath) + ".html";
            string outputPath = Path.Combine(outputDir, outputFileName);

            doc.Save(outputPath, saveOptions);

            if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            {
                throw new InvalidOperationException($"Failed to create HTML output for '{docxPath}'.");
            }
        }

        Console.WriteLine($"Successfully converted {docxFiles.Length} DOCX file(s) to HTML with round‑trip information (if supported).");
    }

    private static void CreateSampleDocx(string filePath, string title, string body)
    {
        if (File.Exists(filePath))
            return;

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln(title);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln(body);
        doc.Save(filePath, SaveFormat.Docx);
    }
}
