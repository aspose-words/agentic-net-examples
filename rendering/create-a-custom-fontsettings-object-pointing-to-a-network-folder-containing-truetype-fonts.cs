using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

namespace FontSettingsExample
{
    public class Program
    {
        public static void Main()
        {
            // Define a UNC path that simulates a network folder containing TrueType fonts.
            // For the purpose of this example we create a local temporary folder and treat its path as a UNC path.
            string localTempFolder = Path.Combine(Path.GetTempPath(), "NetworkFonts");
            Directory.CreateDirectory(localTempFolder);
            // Example UNC path (replace with an actual network share in real scenarios)
            string networkFontFolder = @"\\localhost\" + new DirectoryInfo(localTempFolder).Name;

            // Create a FontSettings object and point it to the network font folder.
            FontSettings fontSettings = new FontSettings();
            // The second parameter indicates whether to scan subfolders.
            fontSettings.SetFontsFolder(networkFontFolder, false);

            // Create a simple document and assign the custom FontSettings.
            Document doc = new Document();
            doc.FontSettings = fontSettings;

            // Add a paragraph that uses a font expected to be found in the network folder.
            // If the font is not present, Aspose.Words will fall back to a default font.
            var builder = new DocumentBuilder(doc);
            builder.Font.Name = "Arial"; // Example font name.
            builder.Writeln("This text is rendered using the custom FontSettings pointing to a network folder.");

            // Save the document to verify that the FontSettings were applied without errors.
            string outputPath = Path.Combine(Path.GetTempPath(), "FontSettingsOutput.docx");
            doc.Save(outputPath);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved successfully.");

            // Clean up temporary resources (optional).
            // File.Delete(outputPath);
            // Directory.Delete(localTempFolder, true);
        }
    }
}
