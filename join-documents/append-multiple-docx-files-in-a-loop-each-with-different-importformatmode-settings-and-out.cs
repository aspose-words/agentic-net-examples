using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create sample source DOCX files.
        // -----------------------------------------------------------------
        var sourceFiles = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(outputDir, $"Source{i}.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the content of source document {i}.");

            // Simple style variation for demonstration.
            if (i == 2)
                builder.Font.Bold = true;
            else if (i == 3)
                builder.Font.Italic = true;

            doc.Save(filePath, SaveFormat.Docx);
            sourceFiles.Add(filePath);
        }

        // -----------------------------------------------------------------
        // Destination document that will hold the combined content.
        // -----------------------------------------------------------------
        var combinedDoc = new Document();

        // Append each source document with a different ImportFormatMode.
        for (int i = 0; i < sourceFiles.Count; i++)
        {
            var srcDoc = new Document(sourceFiles[i]);

            ImportFormatMode mode = i switch
            {
                0 => ImportFormatMode.KeepSourceFormatting,
                1 => ImportFormatMode.UseDestinationStyles,
                _ => ImportFormatMode.KeepDifferentStyles
            };

            combinedDoc.AppendDocument(srcDoc, mode);
        }

        // -----------------------------------------------------------------
        // Validation: ensure that at least the expected number of sections
        // were imported (the destination document starts with one empty section).
        // -----------------------------------------------------------------
        if (combinedDoc.Sections.Count < sourceFiles.Count)
            throw new InvalidOperationException("The combined document does not contain the expected number of sections.");

        // Save the combined document as PDF.
        string pdfPath = Path.Combine(outputDir, "Combined.pdf");
        combinedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the combined PDF file.", pdfPath);
    }
}
