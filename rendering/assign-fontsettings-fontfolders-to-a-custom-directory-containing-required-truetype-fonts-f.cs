using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a custom fonts folder.
        string customFontsDir = Path.Combine(artifactsDir, "CustomFonts");
        Directory.CreateDirectory(customFontsDir);

        // Attempt to copy a system TrueType font into the custom folder.
        // This ensures the folder actually contains a font file for the example.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();
        if (systemFontFolders.Length > 0)
        {
            string firstSystemFolder = systemFontFolders[0];
            string[] ttfFiles = Directory.GetFiles(firstSystemFolder, "*.ttf");
            if (ttfFiles.Length > 0)
            {
                string sourceFont = ttfFiles[0];
                string destFont = Path.Combine(customFontsDir, Path.GetFileName(sourceFont));
                File.Copy(sourceFont, destFont, true);
            }
        }

        // Assign the custom fonts folder to Aspose.Words font settings.
        // The second argument 'true' enables recursive scanning of subfolders.
        FontSettings.DefaultInstance.SetFontsFolder(customFontsDir, true);

        // Create a simple document that uses a font likely present in the copied file.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Use the name of the copied font if known; otherwise fall back to a common font.
        builder.Font.Name = "Arial";
        builder.Writeln("This document is rendered using a custom font folder.");

        // Render the document to PDF.
        string pdfPath = Path.Combine(artifactsDir, "RenderedDocument.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        // Optionally, clean up: reset font sources to original state.
        FontSettings.DefaultInstance.ResetFontSources();
    }
}
