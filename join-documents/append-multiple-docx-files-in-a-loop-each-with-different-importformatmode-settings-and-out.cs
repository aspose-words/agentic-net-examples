using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder to hold sample source documents
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        Directory.CreateDirectory(inputFolder);

        // Create sample DOCX files with simple content
        string[] sourceFiles = new string[3];
        for (int i = 0; i < sourceFiles.Length; i++)
        {
            string filePath = Path.Combine(inputFolder, $"Sample{i + 1}.docx");
            CreateSampleDocument(filePath, $"This is the content of document {i + 1}.");
            sourceFiles[i] = filePath;
        }

        // Define ImportFormatMode settings for each document
        ImportFormatMode[] importModes = new ImportFormatMode[]
        {
            ImportFormatMode.KeepSourceFormatting,
            ImportFormatMode.UseDestinationStyles,
            ImportFormatMode.KeepDifferentStyles
        };

        // Load the first document as the base of the combined document
        Document combinedDoc = new Document(sourceFiles[0]);

        // Append remaining documents using the specified ImportFormatMode
        for (int i = 1; i < sourceFiles.Length; i++)
        {
            Document srcDoc = new Document(sourceFiles[i]);
            combinedDoc.AppendDocument(srcDoc, importModes[i]);
        }

        // Save the combined document as PDF
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "Combined.pdf");
        combinedDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validation: ensure the PDF file exists and contains expected number of sections
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("The combined PDF file was not created.");
        }

        // The combined document should have as many sections as source documents
        if (combinedDoc.Sections.Count != sourceFiles.Length)
        {
            throw new InvalidOperationException($"Expected {sourceFiles.Length} sections, but found {combinedDoc.Sections.Count}.");
        }
    }

    // Helper method to create a simple DOCX file with given text
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);
    }
}
