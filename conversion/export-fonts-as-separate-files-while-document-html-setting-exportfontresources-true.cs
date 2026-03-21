using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ExportFontsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the folder where output will be written.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Create a simple document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Arial";
            builder.Writeln("Hello, world! This document demonstrates exporting fonts.");

            // Configure HTML save options to export each used font as a separate file.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,
                FontsFolder = Path.Combine(artifactsDir, "Fonts")
            };

            // Ensure the fonts folder exists.
            Directory.CreateDirectory(saveOptions.FontsFolder);

            // Save the document as HTML. Fonts will be written to the specified folder.
            string htmlPath = Path.Combine(artifactsDir, "Rendering.html");
            doc.Save(htmlPath, saveOptions);

            Console.WriteLine($"HTML saved to: {htmlPath}");
            Console.WriteLine($"Fonts saved to: {saveOptions.FontsFolder}");
        }
    }
}
