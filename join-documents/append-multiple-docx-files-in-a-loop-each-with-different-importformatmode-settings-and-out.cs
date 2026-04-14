using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder to hold temporary source documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        Directory.CreateDirectory(inputFolder);

        // Create three sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            Document srcDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(srcDoc);
            builder.Writeln($"This is the content of source document {i}.");
            string srcPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            srcDoc.Save(srcPath, SaveFormat.Docx);
        }

        // Destination document that will receive the appended content.
        Document dstDoc = new Document();

        // Different ImportFormatMode values to use for each append.
        ImportFormatMode[] modes = new[]
        {
            ImportFormatMode.UseDestinationStyles,
            ImportFormatMode.KeepSourceFormatting,
            ImportFormatMode.KeepDifferentStyles
        };

        // Append each source document with a distinct ImportFormatMode.
        for (int i = 1; i <= 3; i++)
        {
            string srcPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            Document srcDoc = new Document(srcPath);
            ImportFormatMode mode = modes[(i - 1) % modes.Length];
            dstDoc.AppendDocument(srcDoc, mode);
        }

        // Save the combined document as PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "Combined.pdf");
        dstDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The combined PDF file was not created.");

        // Optional: indicate success.
        Console.WriteLine($"Combined PDF created at: {outputPdf}");
    }
}
