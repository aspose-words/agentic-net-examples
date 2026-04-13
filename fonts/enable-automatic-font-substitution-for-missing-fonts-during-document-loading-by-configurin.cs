using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path for the temporary document that uses a missing font.
        string sourceDocPath = Path.Combine(artifactsDir, "MissingFont.docx");

        // Create a sample document with a font that is unlikely to exist on the system.
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Font.Name = "MissingFontXYZ";
        builder.Writeln("This paragraph is formatted with a missing font.");
        tempDoc.Save(sourceDocPath);

        // Configure FontSettings to substitute missing fonts with a known font (e.g., Arial).
        FontSettings fontSettings = new FontSettings();
        // Use the default font substitution rule.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
        // Enable additional font info substitution for better matching (optional).
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // Apply the FontSettings to LoadOptions.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = fontSettings;

        // Load the document using the configured LoadOptions.
        Document loadedDoc = new Document(sourceDocPath, loadOptions);

        // Keep original font metrics after substitution (optional, default is true).
        loadedDoc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Save the result to PDF to trigger font substitution.
        string resultPath = Path.Combine(artifactsDir, "Result.pdf");
        loadedDoc.Save(resultPath);

        // Simple verification that the output file was created.
        if (File.Exists(resultPath))
        {
            Console.WriteLine("Document saved successfully: " + resultPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
