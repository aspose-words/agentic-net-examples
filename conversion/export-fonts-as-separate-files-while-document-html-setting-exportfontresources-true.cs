using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ExportFontsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the input document and where output will be written.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Load the source document.
            Document doc = new Document(Path.Combine(artifactsDir, "Rendering.docx"));

            // Configure HTML save options to export each used font as a separate file.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,                     // Enable font exporting.
                FontsFolder = Path.Combine(artifactsDir, "Fonts") // Folder where the font files will be saved.
            };

            // Ensure the fonts folder exists; Aspose.Words will also create it if missing.
            Directory.CreateDirectory(saveOptions.FontsFolder);

            // Save the document as HTML. Fonts will be written to the specified folder.
            doc.Save(Path.Combine(artifactsDir, "Rendering.html"), saveOptions);
        }
    }
}
