using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfToHighResPngConverter
{
    /// <summary>
    /// Converts each page of a PDF document that contains vector graphics to a high‑resolution PNG image.
    /// </summary>
    /// <param name="pdfPath">Full path to the source PDF file.</param>
    /// <param name="outputFolder">Folder where the PNG files will be saved.</param>
    /// <param name="dpi">Desired resolution in dots per inch (e.g., 300 or higher for high fidelity).</param>
    public static void Convert(string pdfPath, string outputFolder, int dpi = 300)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document. No special load options are required for this scenario.
        Document doc = new Document(pdfPath);

        // Configure image save options for PNG output.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Set the resolution (both horizontal and vertical) to the desired DPI.
            Resolution = dpi,

            // Enable high‑quality rendering algorithms (slower but better visual fidelity).
            UseHighQualityRendering = true,

            // Enable anti‑aliasing to smooth edges of vector graphics.
            UseAntiAliasing = true
        };

        // Render each page individually.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Restrict rendering to the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name (e.g., Page_1.png, Page_2.png, ...).
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the page as a PNG image using the configured options.
            doc.Save(outputPath, pngOptions);
        }
    }
}

public class Program
{
    /// <summary>
    /// Entry point required for a console application. Demonstrates usage of the converter.
    /// </summary>
    /// <param name="args">[0] = PDF path, [1] = output folder, [2] (optional) = DPI.</param>
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <pdfPath> <outputFolder> [dpi]");
            return;
        }

        string pdfPath = args[0];
        string outputFolder = args[1];
        int dpi = 300;
        if (args.Length >= 3 && int.TryParse(args[2], out int parsedDpi))
            dpi = parsedDpi;

        try
        {
            PdfToHighResPngConverter.Convert(pdfPath, outputFolder, dpi);
            Console.WriteLine($"Conversion completed. PNG files saved to '{outputFolder}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
