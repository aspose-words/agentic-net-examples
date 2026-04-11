using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple DOCX document.
        string docxPath = Path.Combine(artifactsDir, "Sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Writeln("Hello, this is a sample document created for MHTML conversion.");
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the document (demonstrates the load rule).
        Document loadedDoc = new Document(docxPath);

        // Configure MHTML save options to embed CSS (and other resources) directly.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            CssStyleSheetType = CssStyleSheetType.Embedded, // Embed CSS into the MHTML.
            ExportImagesAsBase64 = true,                     // Embed images.
            ExportFontResources = true,                      // Export fonts.
            ExportFontsAsBase64 = true,                      // Embed fonts.
            PrettyFormat = true
        };

        // Save as MHTML.
        string mhtmlPath = Path.Combine(artifactsDir, "Sample.mht");
        loadedDoc.Save(mhtmlPath, saveOptions);

        // Validate that the output file exists and is not empty.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
        {
            throw new InvalidOperationException("MHTML conversion failed: output file is missing or empty.");
        }

        // Verify that CSS is embedded by checking for a <style> tag.
        string mhtmlContent = File.ReadAllText(mhtmlPath);
        if (!mhtmlContent.Contains("<style"))
        {
            throw new InvalidOperationException("CSS was not embedded in the MHTML output.");
        }

        // Indicate successful conversion.
        Console.WriteLine("DOCX successfully converted to MHTML with embedded CSS.");
    }
}
