using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputDocPath = "/path/to/input/document.doc";

        // Path where the rendered PDF will be saved.
        string outputPdfPath = "/path/to/output/document.pdf";

        // Linux font folder that contains TrueType fonts.
        // Adjust this path if your fonts are located elsewhere.
        string linuxFontsFolder = "/usr/share/fonts";

        // Load the DOC file.
        Document doc = new Document(inputDocPath);

        // Create a FontSettings instance and point it to the Linux fonts folder.
        // The second argument (true) makes the search recursive, including sub‑folders.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(linuxFontsFolder, true);

        // Assign the custom FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to PDF (or any other fixed‑page format).
        // Using the overload that specifies the SaveFormat ensures correct rendering.
        doc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
