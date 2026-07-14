using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder to hold the sample source DOCX files.
        string sourceFolder = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
        Directory.CreateDirectory(sourceFolder);

        // Create a few sample DOCX documents programmatically.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is the content of sample document {i}.");
            string docPath = Path.Combine(sourceFolder, $"Doc{i}.docx");
            sampleDoc.Save(docPath, SaveFormat.Docx);
        }

        // Create the master document that will receive all appended content.
        Document masterDoc = new Document();

        // Append each DOCX file in the folder to the master document.
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            Document srcDoc = new Document(filePath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the combined document as a PDF file.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Optional: inform the user that the process completed.
        Console.WriteLine($"Merged PDF created at: {outputPdf}");
    }
}
