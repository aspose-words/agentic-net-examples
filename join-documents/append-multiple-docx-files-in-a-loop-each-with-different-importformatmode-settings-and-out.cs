using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample source DOCX files
        List<string> sourcePaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string srcPath = Path.Combine(outputDir, $"Source{i}.docx");
            Document srcDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(srcDoc);
            builder.Writeln($"This is the content of source document {i}.");
            srcDoc.Save(srcPath, SaveFormat.Docx);
            sourcePaths.Add(srcPath);
        }

        // Define different ImportFormatMode values to use
        ImportFormatMode[] importModes = new[]
        {
            ImportFormatMode.KeepSourceFormatting,
            ImportFormatMode.UseDestinationStyles,
            ImportFormatMode.KeepDifferentStyles
        };

        // Destination document that will hold the combined content
        Document destDoc = new Document();

        // Append each source document using a different ImportFormatMode
        for (int i = 0; i < sourcePaths.Count; i++)
        {
            Document src = new Document(sourcePaths[i]);
            ImportFormatMode mode = importModes[i % importModes.Length];
            destDoc.AppendDocument(src, mode);
        }

        // Save the combined document as PDF
        string pdfPath = Path.Combine(outputDir, "Combined.pdf");
        destDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validation: ensure PDF was created
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The combined PDF file was not created.");
        }

        // Validation: ensure the combined document contains sections from all sources
        if (destDoc.Sections.Count < sourcePaths.Count)
        {
            throw new InvalidOperationException("The combined document does not contain all source sections.");
        }
    }
}
