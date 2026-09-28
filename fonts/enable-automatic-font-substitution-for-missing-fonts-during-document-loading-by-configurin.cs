using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare directories and file names
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeFontDemo");
        Directory.CreateDirectory(baseDir);
        string originalPath = Path.Combine(baseDir, "original.docx");
        string substitutedPath = Path.Combine(baseDir, "substituted.docx");

        // Create a document that uses a non‑existent font
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "NonExistentFont";
        builder.Writeln("This text uses a font that is not installed on the system.");
        doc.Save(originalPath);

        // Configure FontSettings for automatic substitution
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = true;
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Load the document and apply the FontSettings
        Document loadedDoc = new Document(originalPath);
        loadedDoc.FontSettings = fontSettings;

        // Save the document after substitution
        loadedDoc.Save(substitutedPath);

        // Validate that the font was substituted
        Run firstRun = (Run)loadedDoc.GetChild(NodeType.Run, 0, true);
        string substitutedFontName = firstRun.Font.Name ?? string.Empty;
        bool outputExists = File.Exists(substitutedPath);

        var result = new
        {
            OriginalFont = "NonExistentFont",
            SubstitutedFont = substitutedFontName,
            OutputFile = substitutedPath,
            OutputExists = outputExists
        };

        // Output validation result as JSON
        string jsonResult = JsonConvert.SerializeObject(result, Formatting.Indented);
        Console.WriteLine(jsonResult);
    }
}
