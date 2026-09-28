using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Output PDF path
        const string outputPath = "RenderedDocument.pdf";

        // Create a new blank document
        Document doc = new Document();

        // Add a paragraph with text that uses a font unlikely to be installed
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Sample text with a missing font.");
        run.Font.Name = "NonExistentFont123"; // Non‑existent font
        para.AppendChild(run);
        doc.FirstSection.Body.AppendChild(para);

        // Configure font settings to disable automatic font substitution.
        // The SubstitutionSettings.Enabled property may not exist in some versions,
        // so we set it via reflection when available.
        FontSettings fontSettings = new FontSettings();
        var substitutionSettings = fontSettings.SubstitutionSettings;
        var enabledProp = substitutionSettings.GetType().GetProperty("Enabled");
        if (enabledProp != null && enabledProp.CanWrite)
        {
            enabledProp.SetValue(substitutionSettings, false);
        }
        doc.FontSettings = fontSettings;

        // Render the document to PDF
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF file was created
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the rendered PDF at '{outputPath}'.");
        }

        Console.WriteLine($"Document rendered successfully to '{outputPath}'.");
    }
}
