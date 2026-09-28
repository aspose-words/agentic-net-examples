using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Simulate a USB drive folder for fonts.
        string usbFontsPath = Path.Combine(Path.GetTempPath(), "UsbFonts");
        Directory.CreateDirectory(usbFontsPath);

        // Optionally copy a system font to the simulated USB folder if it exists.
        // This step is not required for the example to run, but demonstrates loading a real font.
        try
        {
            string systemFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
            if (File.Exists(systemFontPath))
            {
                string destFontPath = Path.Combine(usbFontsPath, "arial.ttf");
                File.Copy(systemFontPath, destFontPath, true);
            }
        }
        catch
        {
            // Ignore any errors copying the font; the example will still demonstrate FontSettings usage.
        }

        // Configure FontSettings to use the USB fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(usbFontsPath, false);

        // Create a sample document containing special symbols.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("Special symbols: Ω, 漢字, 😀, 𝄞");

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Render the document to PDF.
        string outputPath = Path.Combine(Path.GetTempPath(), "Rendered.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the rendered PDF file.");

        Console.WriteLine($"Document rendered successfully to: {outputPath}");
    }
}
