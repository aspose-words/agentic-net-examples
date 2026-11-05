using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a font that may be missing on the system.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph is initially set to Times New Roman.");

        // Define a custom font substitution rule: use Calibri when the requested font is unavailable.
        FontSettings fontSettings = new FontSettings();
        // When a font cannot be found, Aspose.Words will fall back to the default font specified here.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Calibri";

        // Assign the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        bool fileExists = File.Exists(outputPath);

        // Output verification results as JSON.
        var verificationInfo = new
        {
            OutputFileExists = fileExists,
            DefaultFontSubstitution = fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName
        };

        string json = JsonConvert.SerializeObject(verificationInfo, Formatting.Indented);
        Console.WriteLine(json);
    }
}
