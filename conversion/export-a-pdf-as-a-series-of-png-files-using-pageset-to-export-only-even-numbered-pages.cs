using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfEvenPagesToPng
{
    public static void Main()
    {
        // Define folders for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page document.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Add six pages with simple text.
        for (int i = 1; i <= 6; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 6)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF – this will be the source PDF.
        string pdfPath = Path.Combine(artifactsDir, "Sample.pdf");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF back into Aspose.Words.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Prepare image save options to export only even‑numbered pages.
        // -----------------------------------------------------------------
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Use the built‑in PageSet.Even to filter pages.
            PageSet = PageSet.Even,

            // The callback will create a separate PNG file for each page.
            PageSavingCallback = new EvenPageSavingCallback(artifactsDir)
        };

        // The file name supplied here is ignored because the callback sets the name.
        pdfDoc.Save(Path.Combine(artifactsDir, "Ignored.png"), pngOptions);

        // -----------------------------------------------------------------
        // 4. Verify that PNG files for even pages were created.
        // -----------------------------------------------------------------
        string[] pngFiles = Directory.GetFiles(artifactsDir, "Page_*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were generated.");

        Console.WriteLine($"Generated {pngFiles.Length} PNG file(s) for even pages:");
        foreach (string file in pngFiles)
            Console.WriteLine($" - {Path.GetFileName(file)}");
    }

    // Callback that saves each processed page to a separate PNG file.
    private class EvenPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputDir;

        public EvenPageSavingCallback(string outputDir)
        {
            _outputDir = outputDir;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // PageIndex is zero‑based; add 1 for a human‑readable number.
            string fileName = $"Page_{args.PageIndex + 1}.png";
            args.PageFileName = Path.Combine(_outputDir, fileName);
            // Keep the default behavior of closing the stream after writing.
            args.KeepPageStreamOpen = false;
        }
    }
}
