using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Directory for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Destination document that will hold all appended content.
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Combined Document Start");
        dstBuilder.Writeln(); // Add a blank line.

        // Different ImportFormatMode values to use for each source document.
        List<ImportFormatMode> importModes = new List<ImportFormatMode>
        {
            ImportFormatMode.UseDestinationStyles,
            ImportFormatMode.KeepSourceFormatting,
            ImportFormatMode.KeepDifferentStyles
        };

        // Append a source document for each mode.
        for (int i = 0; i < importModes.Count; i++)
        {
            // Create a simple source document with unique text.
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            srcBuilder.Writeln($"Source Document {i + 1}");
            srcBuilder.Writeln($"This document is appended using {importModes[i]} mode.");
            srcBuilder.Writeln(); // Separate sections.

            // Append the source document to the destination using the current mode.
            dstDoc.AppendDocument(srcDoc, importModes[i]);
        }

        // Save the combined document as PDF.
        string outputPdfPath = Path.Combine(artifactsDir, "Combined.pdf");
        dstDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created and contains at least one page.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The combined PDF file was not created.");
        }

        Document pdfDoc = new Document(outputPdfPath);
        if (pdfDoc.PageCount == 0)
        {
            throw new InvalidOperationException("The combined PDF file contains no pages.");
        }

        // Optional: indicate success (no interactive output required).
        // The program will exit normally if no exception is thrown.
    }
}
