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
            // Define a folder path that simulates a network share containing TrueType fonts.
            // In a real scenario this would be a UNC path like \\server\share\fonts.
            string networkFontsFolder = Path.Combine(Path.GetTempPath(), "NetworkFonts");
            Directory.CreateDirectory(networkFontsFolder);

            // Attempt to copy a system TrueType font into the folder for demonstration purposes.
            // This step is optional and may be skipped if the source font file does not exist.
            try
            {
                string systemFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
                if (File.Exists(systemFontPath))
                {
                    string destFontPath = Path.Combine(networkFontsFolder, "arial.ttf");
                    File.Copy(systemFontPath, destFontPath, true);
                }
            }
            catch
            {
                // Ignore any errors copying the font; the example will still run.
            }

            // Create a custom FontSettings instance and point it to the network fonts folder.
            FontSettings fontSettings = new FontSettings();
            // The second parameter indicates whether to scan subfolders (false = only the specified folder).
            fontSettings.SetFontsFolder(networkFontsFolder, false);

            // Create a simple document and assign the custom FontSettings.
            Document doc = new Document();
            doc.FontSettings = fontSettings;

            // Add a paragraph that uses a font expected to be found in the custom folder.
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            Run run = new Run(doc, "This text is rendered using a font from the custom network folder.");
            // Optionally specify the font name; if the font is not found, Aspose.Words will fall back.
            run.Font.Name = "Arial";
            para.AppendChild(run);

            // Render the document to PDF.
            string outputPath = Path.Combine(Path.GetTempPath(), "CustomFontSettingsOutput.pdf");
            doc.Save(outputPath, SaveFormat.Pdf);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The PDF file was not created.");

            // Example completed successfully.
            Console.WriteLine("Document saved to: " + outputPath);
        }
    }
}
