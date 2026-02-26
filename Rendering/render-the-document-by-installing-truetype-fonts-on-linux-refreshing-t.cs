using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class RenderDocumentWithLinuxFonts
{
    static void Main()
    {
        // Path to the folder that contains the TrueType fonts you want to use.
        // This folder can be anywhere on the file system; we will copy its contents
        // to a system font directory on Linux.
        string sourceFontsFolder = Path.Combine(Directory.GetCurrentDirectory(), "MyFonts");

        // Target system font directory on Linux. Adjust if needed.
        string systemFontsFolder = "/usr/share/fonts/truetype/custom";

        // Ensure we are running on Linux before attempting to install fonts.
        if (Environment.OSVersion.Platform == PlatformID.Unix)
        {
            // Create the target directory if it does not exist.
            if (!Directory.Exists(systemFontsFolder))
                Directory.CreateDirectory(systemFontsFolder);

            // Copy all .ttf files from the source folder to the system fonts folder.
            foreach (string fontFile in Directory.GetFiles(sourceFontsFolder, "*.ttf"))
            {
                string destFile = Path.Combine(systemFontsFolder, Path.GetFileName(fontFile));
                File.Copy(fontFile, destFile, overwrite: true);
            }

            // Refresh the font cache so the newly copied fonts become visible to the OS.
            // The command "fc-cache -f -v" forces a rebuild of the font cache.
            var fcCache = new ProcessStartInfo
            {
                FileName = "fc-cache",
                Arguments = "-f -v",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using (var proc = Process.Start(fcCache))
            {
                proc.WaitForExit();
                // Optionally read output for debugging:
                // string output = proc.StandardOutput.ReadToEnd();
            }
        }

        // Configure Aspose.Words to look for fonts in the folder we just populated.
        FontSettings fontSettings = new FontSettings();
        // The second argument (recursive) tells Aspose.Words to search subfolders as well.
        fontSettings.SetFontsFolder(systemFontsFolder, recursive: true);

        // Load the document with the custom FontSettings.
        // Replace "input.docx" with the path to your source document.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        LoadOptions loadOptions = new LoadOptions
        {
            FontSettings = fontSettings
        };
        Document doc = new Document(inputPath, loadOptions);

        // Render the document to PDF (or any other fixed‑page format).
        // Replace "output.pdf" with the desired output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.pdf");
        doc.Save(outputPath);
    }
}
