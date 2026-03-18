using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Fonts;

class InstallFontsOnLinux
{
    static void Main()
    {
        // Paths – adjust as needed.
        string sourceFontsDir = "/usr/share/fonts/truetype/custom"; // Directory containing .ttf files to install.
        string targetFontsDir = "/app/fonts";                       // Directory where Aspose.Words will look for fonts.

        // Ensure the target directory exists.
        if (!Directory.Exists(targetFontsDir))
            Directory.CreateDirectory(targetFontsDir);

        // Copy all TrueType font files (*.ttf) from the source to the target directory.
        foreach (string fontFilePath in Directory.GetFiles(sourceFontsDir, "*.ttf", SearchOption.AllDirectories))
        {
            string destPath = Path.Combine(targetFontsDir, Path.GetFileName(fontFilePath));
            // Overwrite if the file already exists.
            File.Copy(fontFilePath, destPath, true);
        }

        // Configure Aspose.Words to use the target fonts folder.
        FontSettings fontSettings = new FontSettings();
        // 'true' enables recursive scanning of subfolders (if any).
        fontSettings.SetFontsFolder(targetFontsDir, true);

        // Load a document with the custom FontSettings.
        // Replace "input.docx" with the path to your source document.
        string inputPath = "input.docx";
        LoadOptions loadOptions = new LoadOptions
        {
            FontSettings = fontSettings
        };
        Document doc = new Document(inputPath, loadOptions);

        // Optionally, embed the fonts into the saved document.
        doc.FontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;
        doc.FontInfos.EmbedTrueTypeFonts = true;
        doc.FontInfos.EmbedSystemFonts = true;
        doc.FontInfos.SaveSubsetFonts = true;

        // Save the processed document.
        // Replace "output.pdf" with the desired output path and format.
        string outputPath = "output.pdf";
        doc.Save(outputPath);
    }
}
