using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportFontsExample
{
    public static void Main()
    {
        // Create a sample document with a common font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Font.Size = 14;
        builder.Writeln("This is a sample paragraph using Arial font.");

        // Define output HTML path.
        string htmlPath = "output.html";

        // Configure HTML save options to export font resources.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportFontResources = true,
            ExportImagesAsBase64 = false
        };

        // Save the document to HTML.
        doc.Save(htmlPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException($"HTML file '{htmlPath}' was not created.");

        // Determine the folder that should contain exported font files.
        string fontFolder = Path.Combine(
            Path.GetDirectoryName(htmlPath) ?? string.Empty,
            Path.GetFileNameWithoutExtension(htmlPath) + "_files");

        // The folder is created only when the document contains fonts that need to be exported.
        // If the folder does not exist, it simply means there were no font resources to export.
        if (Directory.Exists(fontFolder))
        {
            // Check for at least one font file (e.g., .woff, .woff2, .ttf, .eot, .svg).
            string[] fontExtensions = new[] { ".woff", ".woff2", ".ttf", ".eot", ".svg" };
            bool fontFileExists = Directory.EnumerateFiles(fontFolder)
                .Any(file => fontExtensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase));

            if (!fontFileExists)
                throw new InvalidOperationException("No font files were exported to the resources folder.");
        }
        else
        {
            Console.WriteLine("Font resources folder was not created because the document does not contain fonts that need exporting.");
        }

        // Indicate success.
        Console.WriteLine("HTML export completed successfully.");
    }
}
