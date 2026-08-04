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

        // Create a temporary document that uses a font that likely does not exist.
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Font.Name = "MissingFontXYZ";
        builder.Writeln("This text uses a missing font and will be substituted.");

        // Save the temporary document.
        string tempPath = Path.Combine(artifactsDir, "Temp.docx");
        tempDoc.Save(tempPath);

        // Configure FontSettings to substitute missing fonts with a known font (Arial).
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Load the document with the configured FontSettings.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = fontSettings;
        Document loadedDoc = new Document(tempPath, loadOptions);

        // Save the loaded document to PDF; missing fonts will be substituted automatically.
        string outputPath = Path.Combine(artifactsDir, "Result.pdf");
        loadedDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
