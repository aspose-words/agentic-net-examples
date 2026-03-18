using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input PDF file path
        string pdfPath = "input.pdf";

        // Output PNG file pattern (Aspose.Words will append page numbers for multi‑page documents)
        string pngPath = "output.png";

        // Load the PDF document (no special load options required for this scenario)
        Document doc = new Document(pdfPath);

        // Configure image save options for high‑resolution PNG output
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render at 300 DPI for detailed analysis
            Resolution = 300,

            // Enable high‑quality rendering and anti‑aliasing for sharper results
            UseHighQualityRendering = true,
            UseAntiAliasing = true,

            // By default all pages are rendered; you can set a specific PageSet if needed
            // PageSet = PageSet.All
        };

        // Save each page of the PDF as a separate PNG image
        doc.Save(pngPath, pngOptions);
    }
}
