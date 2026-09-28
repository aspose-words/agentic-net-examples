using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Configure font settings to use the system fonts folder.
        FontSettings fontSettings = new FontSettings();
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        fontSettings.SetFontsFolder(fontsFolder, false);
        doc.FontSettings = fontSettings;

        // Add a paragraph containing text with ligatures (fi, fl, ffi).
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "office affinity file");
        run.Font.Name = "Calibri"; // Calibri supports OpenType ligatures.
        run.Font.Size = 24;
        paragraph.AppendChild(run);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Render the document to PDF.
        string outputPath = "Output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"PDF file was not created at '{outputPath}'.");
        }
    }
}
