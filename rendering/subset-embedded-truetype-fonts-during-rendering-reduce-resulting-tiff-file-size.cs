using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a simple document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Enable embedding of TrueType fonts and request subsetting, if the FontInfos collection is available.
        FontInfoCollection fontInfos = doc.FontInfos;
        if (fontInfos != null)
        {
            fontInfos.EmbedTrueTypeFonts = true;   // Embed TrueType fonts.
            fontInfos.SaveSubsetFonts = true;      // Save only the used glyphs.
        }

        // Ensure the layout is up‑to‑date before accessing PageCount.
        doc.UpdatePageLayout();

        // Configure TIFF rendering options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300 // Render at 300 DPI.
        };

        // Render each page of the document to a separate TIFF file.
        for (int i = 0; i < doc.PageCount; i++)
        {
            options.PageSet = new PageSet(i); // Render only the current page.
            string outputPath = $"Page_{i + 1}.tiff";
            doc.Save(outputPath, options);
        }
    }
}
