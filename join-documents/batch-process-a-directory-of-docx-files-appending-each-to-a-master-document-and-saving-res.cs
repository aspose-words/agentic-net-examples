using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder to hold the sample source DOCX files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        Directory.CreateDirectory(inputFolder);

        // Create a few sample DOCX documents inside the folder.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample document {i} content.");
            string samplePath = Path.Combine(inputFolder, $"Doc{i}.docx");
            sampleDoc.Save(samplePath, SaveFormat.Docx);
        }

        // Create the master document that will receive the appended content.
        Document masterDoc = new Document();

        // Append each DOCX file from the folder to the master document.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string filePath in docxFiles)
        {
            Document srcDoc = new Document(filePath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged result as a PDF file.
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "MasterDocument.pdf");
        masterDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Validation: ensure the merged document contains content from all source files.
        Document verificationDoc = new Document(outputPdfPath);
        string mergedText = verificationDoc.GetText();

        for (int i = 1; i <= 3; i++)
        {
            if (!mergedText.Contains($"Sample document {i} content."))
            {
                throw new InvalidOperationException($"Merged content missing for source document {i}.");
            }
        }

        // Cleanup: optional removal of temporary files (comment out if inspection is needed).
        // Directory.Delete(inputFolder, true);
    }
}
