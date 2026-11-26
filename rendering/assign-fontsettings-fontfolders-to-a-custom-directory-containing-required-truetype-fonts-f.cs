using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample document and the rendered PDF.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "RenderedDocument.pdf");

        // Create a custom folder that will hold TrueType fonts.
        string customFontFolder = Path.Combine(Directory.GetCurrentDirectory(), "CustomFonts");
        Directory.CreateDirectory(customFontFolder);

        // OPTIONAL: If a TrueType font file is available on the machine, copy it into the custom folder.
        // This step is safe even if the source font does not exist; the example focuses on assigning the folder.
        try
        {
            string systemFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
            if (File.Exists(systemFontPath))
            {
                string destFontPath = Path.Combine(customFontFolder, "arial.ttf");
                File.Copy(systemFontPath, destFontPath, true);
            }
        }
        catch
        {
            // Ignore any errors while copying system fonts; the folder assignment is still demonstrated.
        }

        // Create a FontSettings instance and assign the custom font folder.
        FontSettings fontSettings = new FontSettings();
        // Assign the folder (recursive search enabled) to the FontSettings.
        fontSettings.SetFontsFolder(customFontFolder, true);

        // Create a new blank document.
        Document doc = new Document();

        // Apply the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Add a paragraph that uses a font expected to be found in the custom folder.
        // If the font is not present, Aspose.Words will fall back to a default font.
        var builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Example font name; adjust if a different custom font is placed in the folder.
        builder.Writeln("This paragraph is rendered using fonts from the custom folder.");

        // Render the document to PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Failed to create the rendered PDF file.");
        }

        // Optionally, output the path of the generated PDF (no interactive prompts).
        Console.WriteLine($"PDF rendered successfully to: {pdfPath}");
    }
}
