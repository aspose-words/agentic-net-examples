using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Simulate a USB drive by creating a temporary folder that will hold font files.
        string usbFontsFolder = Path.Combine(Path.GetTempPath(), "UsbFonts");
        Directory.CreateDirectory(usbFontsFolder);

        // Optionally copy a known system font into the simulated USB folder.
        // This step is safe even if the source font does not exist on the current platform.
        string systemFontsPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        string sourceFont = Path.Combine(systemFontsPath, "arial.ttf");
        string destFont = Path.Combine(usbFontsFolder, "arial.ttf");
        if (File.Exists(sourceFont))
        {
            File.Copy(sourceFont, destFont, true);
        }

        // Configure FontSettings to load fonts from the USB folder.
        FontSettings fontSettings = new FontSettings();
        // The second parameter (false) tells Aspose.Words not to search subfolders.
        fontSettings.SetFontsFolder(usbFontsFolder, false);

        // Create a sample document containing special symbols.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Use a font that is expected to be in the USB folder.
        builder.Writeln("Special symbols: Ω ∑ € 漢字 😀");

        // Apply the custom FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPath = Path.Combine(Path.GetTempPath(), "RenderedDocument.pdf");
        doc.Save(outputPath);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF file was not created.");
        }

        // Optionally, output the location of the generated file (no user interaction required).
        Console.WriteLine($"Document rendered successfully to: {outputPath}");
    }
}
