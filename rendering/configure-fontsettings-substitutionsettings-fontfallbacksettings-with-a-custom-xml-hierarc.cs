using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Directories for temporary files.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for fallback settings files.
        string defaultFallbackPath = Path.Combine(outputDir, "DefaultFallback.xml");
        string customFallbackPath = Path.Combine(outputDir, "CustomFallback.xml");
        string pdfPath = Path.Combine(outputDir, "Result.pdf");

        // 1. Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont"; // Font that does not exist in the sources.
        builder.Writeln("Sample text using a missing font. Unicode range 0x0021-0x00FF:");
        builder.Writeln("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
        builder.Writeln("abcdefghijklmnopqrstuvwxyz");
        builder.Writeln("0123456789");

        // 2. Configure FontSettings to use system fonts.
        FontSettings fontSettings = new FontSettings();
        string systemFontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(systemFontsFolder, false);
        doc.FontSettings = fontSettings;

        // 3. Build automatic fallback settings and save them to a default XML file.
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.BuildAutomatic();
        fallback.Save(defaultFallbackPath);

        // 4. Create a custom fallback XML by modifying the default one.
        //    Replace the first fallback font with a different one (e.g., "Times New Roman").
        string defaultXml = File.ReadAllText(defaultFallbackPath);
        string customXml = defaultXml.Replace("<FallbackFonts>", "<FallbackFonts><Font>Times New Roman</Font>")
                                    .Replace("</FallbackFonts>", "</FallbackFonts>");
        File.WriteAllText(customFallbackPath, customXml);

        // 5. Load the custom fallback settings.
        fallback.Load(customFallbackPath);

        // 6. Save the document to PDF using the configured fallback settings.
        doc.Save(pdfPath);

        // 7. Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF output was not created.");

        // 8. Optionally, save the currently loaded fallback settings to another file for inspection.
        string savedAfterLoadPath = Path.Combine(outputDir, "SavedAfterLoad.xml");
        fallback.Save(savedAfterLoadPath);
    }
}
