using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define the folder that will hold the sample source DOCX files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        Directory.CreateDirectory(inputFolder);

        // Create a few sample DOCX files inside the folder.
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), "First document content.");
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), "Second document content.");
        CreateSampleDocument(Path.Combine(inputFolder, "Doc3.docx"), "Third document content.");

        // Create the master document that will receive all appended documents.
        var masterDoc = new Document();

        // Append each DOCX file found in the input folder to the master document.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            var srcDoc = new Document(filePath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Define the output PDF file path.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");

        // Save the merged document as PDF.
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Optional: indicate successful completion.
        Console.WriteLine($"Merged PDF created at: {outputPdf}");
    }

    // Helper method to create a simple DOCX file with given text.
    private static void CreateSampleDocument(string filePath, string content)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath, SaveFormat.Docx);
    }
}
