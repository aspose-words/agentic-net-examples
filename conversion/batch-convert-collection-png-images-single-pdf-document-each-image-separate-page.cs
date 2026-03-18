using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class Program
{
    static void Main(string[] args)
    {
        // Expected usage: ImageToPdfConverter <output.pdf> <image1.png> [image2.png] ...
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ImageToPdfConverter <output.pdf> <image1.png> [image2.png] ...");
            return;
        }

        string outputPdf = args[0];
        string[] pngFiles = new string[args.Length - 1];
        Array.Copy(args, 1, pngFiles, 0, pngFiles.Length);

        ImageToPdfConverter.ConvertPngsToPdf(pngFiles, outputPdf);
        Console.WriteLine($"PDF created at {outputPdf}");
    }
}

public static class ImageToPdfConverter
{
    /// <summary>
    /// Converts an array of PNG file paths into a single PDF document.
    /// Each image is placed on its own page.
    /// </summary>
    /// <param name="pngFiles">Array of full paths to PNG images.</param>
    /// <param name="outputPdf">Full path for the resulting PDF file.</param>
    public static void ConvertPngsToPdf(string[] pngFiles, string outputPdf)
    {
        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each PNG image and add a page break after it (except after the last image).
        for (int i = 0; i < pngFiles.Length; i++)
        {
            // Insert the image onto the current page.
            builder.InsertImage(pngFiles[i]);

            // Insert a page break to start a new page for the next image.
            if (i < pngFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF using default PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(outputPdf, pdfOptions);
    }
}
