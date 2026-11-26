using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Path for the output TIFF file.
        string outputPath = "OpenTypeLigatures.tiff";

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document uses a font that supports discretionary ligatures (e.g., Gabriola).
        // If the font is not available on the system, Aspose.Words will fall back to a default font.
        Run run = new Run(doc, "office"); // The word "office" contains the "ffi" ligature.
        run.Font.Name = "Gabriola";

        // Add the run to a paragraph and the paragraph to the document.
        Paragraph para = new Paragraph(doc);
        para.AppendChild(run);
        doc.FirstSection.Body.AppendChild(para);

        // Optional: configure FontSettings to point to the system fonts folder.
        // This helps Aspose.Words locate the required font if it is not in the default search locations.
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (Directory.Exists(fontsFolder))
        {
            fontSettings.SetFontsFolder(fontsFolder, false);
        }
        doc.FontSettings = fontSettings;

        // Render the document to a multipage TIFF image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Use default DPI and other settings; they are sufficient for this verification.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Verify that the source document has at least one page.
        if (doc.PageCount < 1)
        {
            throw new InvalidOperationException("The source document does not contain any pages.");
        }

        // Output simple confirmation (no user interaction required).
        Console.WriteLine($"TIFF file generated successfully at '{Path.GetFullPath(outputPath)}'.");
        Console.WriteLine($"Document page count: {doc.PageCount}");
    }
}
