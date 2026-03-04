using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be loaded.
        string inputPath = "MyDir/Document.docx";

        // Optional path where the document will be saved after loading.
        string outputPath = "ArtifactsDir/DocumentWithFonts.docx";

        // Create LoadOptions and assign a FontSettings instance.
        // The default FontSettings already contains a SystemFontSource,
        // which makes Aspose.Words search the operating system's font folders.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = new FontSettings();

        // Load the document using the LoadOptions that contain the font settings.
        Document doc = new Document(inputPath, loadOptions);

        // (Optional) List the font sources that Aspose.Words will use.
        foreach (FontSourceBase source in doc.FontSettings.GetFontsSources())
        {
            Console.WriteLine($"Source type: {source.Type}, Priority: {source.Priority}");
        }

        // Save the document (no modifications needed; this verifies successful loading).
        doc.Save(outputPath);
    }
}
