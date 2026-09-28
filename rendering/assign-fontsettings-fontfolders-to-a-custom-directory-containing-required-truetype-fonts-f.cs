using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder that will hold custom TrueType fonts.
        string customFontFolder = Path.Combine(Path.GetTempPath(), "CustomFonts_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(customFontFolder);

        // Attempt to copy a system TrueType font into the custom folder (if a system font folder exists).
        string systemFontFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        string copiedFontFileName = null;

        if (Directory.Exists(systemFontFolder))
        {
            // Find the first .ttf file in the system fonts directory.
            string systemFontFile = Directory.GetFiles(systemFontFolder, "*.ttf").FirstOrDefault();
            if (!string.IsNullOrEmpty(systemFontFile))
            {
                copiedFontFileName = Path.GetFileName(systemFontFile);
                string destinationPath = Path.Combine(customFontFolder, copiedFontFileName);
                File.Copy(systemFontFile, destinationPath, true);
            }
        }

        // Prepare FontSettings and point it to the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        // Assign the custom folder (recursive search enabled).
        fontSettings.SetFontsFolder(customFontFolder, true);

        // Create a simple document that uses the copied font (if any), otherwise default font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = !string.IsNullOrEmpty(copiedFontFileName)
            ? Path.GetFileNameWithoutExtension(copiedFontFileName) // Use the font name derived from the file.
            : "Times New Roman"; // Fallback to a common font.
        builder.Writeln("This text is rendered using a font loaded from a custom folder.");

        // Apply the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF to trigger font loading.
        string outputPath = Path.Combine(Path.GetTempPath(), "RenderedDocument.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF output file was not created.");

        // (Optional) Clean up temporary resources.
        // Note: In a real scenario you might want to keep the files for inspection.
        // Here we delete the temporary font folder but keep the PDF for verification.
        try
        {
            Directory.Delete(customFontFolder, true);
        }
        catch
        {
            // Ignored – cleanup failure should not affect the example execution.
        }

        // Indicate successful completion (no console interaction required).
    }
}
