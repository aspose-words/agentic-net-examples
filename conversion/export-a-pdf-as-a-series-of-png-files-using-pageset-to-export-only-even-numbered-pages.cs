using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfEvenPagesToPng
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample multi‑page document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 6; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 6)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // 2. Save the document as PDF (the source PDF).
        string pdfPath = Path.Combine(outputDir, "Sample.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // 3. Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // 4. Configure image save options to export only even‑numbered pages.
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // PageSet.Even selects pages with even numbers (2,4,6,…).
            PageSet = PageSet.Even,
            // Use a callback to name each exported PNG file.
            PageSavingCallback = new EvenPageSavingCallback(outputDir)
        };

        // 5. Save the selected pages. The file name supplied here is ignored because
        //    the callback provides explicit names for each page.
        pdfDoc.Save("unused.png", imgOptions);

        // 6. Validate that PNG files for even pages were created.
        string[] pngFiles = Directory.GetFiles(outputDir, "EvenPage_*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were generated for even pages.");

        Console.WriteLine("Exported PNG files:");
        foreach (string file in pngFiles)
            Console.WriteLine(file);
    }

    // Callback that assigns a file name to each page being saved.
    private class EvenPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputFolder;

        public EvenPageSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // PageIndex is zero‑based; add 1 for human‑readable page numbers.
            string fileName = $"EvenPage_{args.PageIndex + 1}.png";
            args.PageFileName = Path.Combine(_outputFolder, fileName);
            // Keep the stream closed after each page is written (default behavior).
            args.KeepPageStreamOpen = false;
        }
    }
}
