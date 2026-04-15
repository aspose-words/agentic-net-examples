using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output and font directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string fontsDir = Path.Combine(artifactsDir, "CustomFonts");
        Directory.CreateDirectory(fontsDir);

        // OPTIONAL: copy a known system font into the custom folder if you want the example to render correctly.
        // This step is safe even if the source font does not exist; the focus is on demonstrating FontSettings usage.
        try
        {
            // Example for Windows: copy Arial.ttf from the system fonts folder.
            string windowsFonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
            if (File.Exists(windowsFonts))
            {
                File.Copy(windowsFonts, Path.Combine(fontsDir, "arial.ttf"), overwrite: true);
            }
        }
        catch
        {
            // Ignore any errors – the example still compiles and runs.
        }

        // Create a simple document that uses a font which may not be installed on the machine.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arvo"; // Font name that might be missing.
        builder.Writeln("Sample text rendered with the custom font.");

        // Assign the custom fonts folder to the document's FontSettings.
        FontSettings fontSettings = new FontSettings();
        // Scan the folder recursively (true) so subfolders are also considered.
        fontSettings.SetFontsFolder(fontsDir, true);
        doc.FontSettings = fontSettings;

        // Save the document to PDF.
        string outputPath = Path.Combine(artifactsDir, "CustomFontDocument.pdf");
        doc.Save(outputPath);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
