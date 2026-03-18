using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class FontSubstitutionExample
{
    static void Main()
    {
        // Paths to the input document and the output file.
        string inputPath = @"C:\Docs\MissingFont.docx";
        string outputPath = @"C:\Docs\Result.pdf";

        // Create a FontSettings instance that will be used during loading.
        FontSettings fontSettings = new FontSettings();

        // Configure the default font substitution rule.
        // If a font used in the document cannot be found, Aspose.Words will replace it with Arial.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Optionally enable other substitution rules (e.g., font info substitution).
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

        // Attach the FontSettings to LoadOptions.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = fontSettings;

        // Load the document using the configured LoadOptions.
        Document doc = new Document(inputPath, loadOptions);

        // Save the document; missing fonts will have been substituted automatically.
        doc.Save(outputPath);
    }
}
