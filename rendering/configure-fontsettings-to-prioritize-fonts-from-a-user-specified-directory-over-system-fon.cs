using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define output and custom fonts directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string customFontsDir = Path.Combine(outputDir, "MyFonts");
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(customFontsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This text is rendered using Arial.");

        // Configure FontSettings to prioritize the custom fonts folder.
        FontSettings fontSettings = new FontSettings();

        // Add the custom folder as the first font source.
        FolderFontSource customFolderSource = new FolderFontSource(customFontsDir, true);
        // Retrieve the existing system font sources.
        FontSourceBase[] systemSources = FontSettings.DefaultInstance.GetFontsSources();

        // Combine custom folder source with the system sources (custom first).
        FontSourceBase[] combinedSources = new FontSourceBase[systemSources.Length + 1];
        combinedSources[0] = customFolderSource;
        Array.Copy(systemSources, 0, combinedSources, 1, systemSources.Length);

        // Apply the combined font sources to the FontSettings instance.
        fontSettings.SetFontsSources(combinedSources);

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");
        doc.Save(pdfPath);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the rendered PDF file.");

        // (Optional) Clean up: restore default font sources if needed.
        // FontSettings.DefaultInstance.SetFontsSources(systemSources);
    }
}
