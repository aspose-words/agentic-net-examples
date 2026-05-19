using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output PNGs.
        string inputFolder = "InputPdfs";
        string outputFolder = "OutputPngs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Step 1: Create a few sample PDF files (the task assumes no external files).
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF document {i}");
            builder.Writeln("This document will be converted to high‑resolution PNG images.");
            // Add a second page to demonstrate multi‑page handling.
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Second page of document {i}.");

            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: {pdfPath}");
        }

        // -----------------------------------------------------------------
        // Step 2: Batch convert each PDF to PNG images at 600 DPI.
        // -----------------------------------------------------------------
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Convert each page of the PDF to a separate PNG file.
            for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
            {
                // Configure image save options: PNG format, 600 DPI, render only the current page.
                ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
                pngOptions.Resolution = 600; // Set both horizontal and vertical resolution.
                pngOptions.PageSet = new PageSet(pageIndex); // Zero‑based page index.

                // Build the output file name: original name + page number.
                string outputFileName = $"{Path.GetFileNameWithoutExtension(pdfFile)}_page{pageIndex + 1}.png";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the page as a PNG image.
                pdfDoc.Save(outputPath, pngOptions);

                // Verify that the PNG file was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException($"Failed to create PNG image: {outputPath}");
            }
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
