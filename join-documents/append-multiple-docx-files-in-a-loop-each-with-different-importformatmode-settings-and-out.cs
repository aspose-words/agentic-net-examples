using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Directory for temporary source documents and final output.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsExample");
        Directory.CreateDirectory(workDir);

        // Create sample source DOCX files.
        string[] sourceFiles = CreateSampleDocuments(workDir);

        // Destination document that will hold the merged content.
        var destination = new Document();

        // Different ImportFormatMode values to use for each append operation.
        ImportFormatMode[] importModes = new[]
        {
            ImportFormatMode.UseDestinationStyles,
            ImportFormatMode.KeepSourceFormatting,
            ImportFormatMode.KeepDifferentStyles
        };

        // Append each source document with a corresponding ImportFormatMode.
        for (int i = 0; i < sourceFiles.Length; i++)
        {
            var srcDoc = new Document(sourceFiles[i]);
            ImportFormatMode mode = importModes[i % importModes.Length];
            destination.AppendDocument(srcDoc, mode);
        }

        // Validate that the expected number of sections are present.
        // Each source document adds one section; the destination started with one.
        int expectedSections = sourceFiles.Length + 1;
        if (destination.Sections.Count != expectedSections)
            throw new InvalidOperationException($"Expected {expectedSections} sections after merging, but found {destination.Sections.Count}.");

        // Save the merged document as PDF.
        string pdfPath = Path.Combine(workDir, "MergedDocument.pdf");
        destination.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("The merged PDF file was not created.", pdfPath);

        // Optional: output the location of the generated PDF.
        Console.WriteLine($"Merged PDF created at: {pdfPath}");
    }

    // Creates a few sample DOCX files with simple text and returns their file paths.
    private static string[] CreateSampleDocuments(string folder)
    {
        var filePaths = new string[3];

        for (int i = 0; i < filePaths.Length; i++)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the content of source document #{i + 1}.");
            string filePath = Path.Combine(folder, $"SourceDoc{i + 1}.docx");
            doc.Save(filePath, SaveFormat.Docx);
            filePaths[i] = filePath;
        }

        return filePaths;
    }
}
