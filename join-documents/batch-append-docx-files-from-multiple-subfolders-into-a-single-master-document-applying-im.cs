using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for the sample input documents.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        // Ensure a clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(baseDir);

        // Create two subfolders, each will contain a couple of DOCX files.
        string[] subFolders = { "FolderA", "FolderB" };
        foreach (string folder in subFolders)
        {
            string folderPath = Path.Combine(baseDir, folder);
            Directory.CreateDirectory(folderPath);

            // Create two sample DOCX files in each subfolder.
            for (int i = 1; i <= 2; i++)
            {
                string fileName = $"Sample_{folder}_{i}.docx";
                string filePath = Path.Combine(folderPath, fileName);

                // Build a simple document with identifiable text.
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"This is {fileName}");
                doc.Save(filePath, SaveFormat.Docx);
            }
        }

        // Create the master document that will receive all appended content.
        Document masterDoc = new Document();
        DocumentBuilder masterBuilder = new DocumentBuilder(masterDoc);
        masterBuilder.Writeln("Master Document - Combined Content");
        masterBuilder.Writeln("-----------------------------------");
        masterBuilder.Writeln();

        // Append every DOCX file found in all subfolders using UseDestinationStyles.
        string[] docFiles = Directory.GetFiles(baseDir, "*.docx", SearchOption.AllDirectories);
        foreach (string docPath in docFiles)
        {
            Document srcDoc = new Document(docPath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
        }

        // Define output PDF path.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");

        // Save the combined document as PDF.
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The merged PDF file was not created.");

        // Additional validation: ensure the master document contains sections from all source docs.
        // Each source document adds at least one section, so total sections should be >= number of source files + 1 (title section).
        int expectedMinSections = docFiles.Length + 1;
        if (masterDoc.Sections.Count < expectedMinSections)
            throw new InvalidOperationException("The merged document does not contain all expected sections.");

        // Indicate successful completion (no interactive output required).
        // The program will exit automatically.
    }
}
