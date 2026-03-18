using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts; // <-- added namespace for FontInfoCollection

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Enable embedding of TrueType fonts and request subsetting.
        FontInfoCollection fontInfos = doc.FontInfos;
        fontInfos.EmbedTrueTypeFonts = true;   // Embed TrueType fonts.
        fontInfos.SaveSubsetFonts = true;      // Save only the used glyphs.

        // Configure TIFF rendering options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300,                     // Render at 300 DPI.
            // ImageSize = new Size(2000, 2600), // Optional: set explicit pixel dimensions.
        };

        // Render each page of the document to a separate TIFF file.
        for (int i = 0; i < doc.PageCount; i++)
        {
            options.PageSet = new PageSet(i);    // Render only the current page.
            string outputPath = $"Page_{i + 1}.tiff";
            doc.Save(outputPath, options);
        }
    }
}
