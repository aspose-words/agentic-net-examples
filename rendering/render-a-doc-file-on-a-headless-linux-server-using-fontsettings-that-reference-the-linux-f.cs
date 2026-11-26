using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;   // Needed for LoadOptions

public class Program
{
    public static void Main()
    {
        // Create a temporary working directory.
        string workingDir = Path.Combine(Path.GetTempPath(), "AsposeRenderExample");
        Directory.CreateDirectory(workingDir);

        // Paths for the intermediate DOCX and final PDF.
        string sourceDocPath = Path.Combine(workingDir, "sample.docx");
        string outputPdfPath = Path.Combine(workingDir, "output.pdf");

        // Configure FontSettings to point to the Linux fonts folder.
        string linuxFontsFolder = "/usr/share/fonts/truetype";
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(linuxFontsFolder, false);

        // Build a simple document that uses a common Linux font.
        Document doc = new Document();
        doc.FontSettings = fontSettings;

        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Hello, Aspose.Words rendering with custom fonts on Linux!");
        run.Font.Name = "DejaVu Sans"; // Typical font available on Linux.
        paragraph.AppendChild(run);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the source DOCX.
        doc.Save(sourceDocPath);

        // Load the document again with the same FontSettings (simulating a separate load on a headless server).
        LoadOptions loadOptions = new LoadOptions { FontSettings = fontSettings };
        Document loadedDoc = new Document(sourceDocPath, loadOptions);

        // Render the document to PDF.
        loadedDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new Exception("Failed to create the PDF output.");
        }

        // Optional cleanup (commented out to allow inspection of generated files).
        // File.Delete(sourceDocPath);
        // File.Delete(outputPdfPath);
        // Directory.Delete(workingDir);
    }
}
