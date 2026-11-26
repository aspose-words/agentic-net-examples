using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input folder and output file paths.
        string inputFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinSample");
        string outputPdfPath = Path.Combine(inputFolder, "Master.pdf");

        // Ensure the input folder exists.
        Directory.CreateDirectory(inputFolder);

        // Create sample DOCX files in the input folder.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is document {i}");
            // Add a page break after each document except the last to separate them.
            if (i < sampleCount)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
            sampleDoc.Save(docPath, SaveFormat.Docx);
        }

        // Create an empty master document.
        Document masterDoc = new Document();

        // Append each DOCX file to the master document.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string filePath in docxFiles)
        {
            Document srcDoc = new Document(filePath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged document as PDF.
        masterDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Validation: ensure content from all source documents is present.
        string mergedText = masterDoc.GetText();
        for (int i = 1; i <= sampleCount; i++)
        {
            if (!mergedText.Contains($"This is document {i}"))
            {
                throw new InvalidOperationException($"Content from source document {i} is missing in the merged result.");
            }
        }
    }
}
