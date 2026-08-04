using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for the demo files.
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeJoinDemo");
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(baseDir);

        // Create two subfolders and seed each with a couple of DOCX files.
        string[] subFolders = { "FolderA", "FolderB" };
        int docCounter = 1;
        foreach (string folder in subFolders)
        {
            string folderPath = Path.Combine(baseDir, folder);
            Directory.CreateDirectory(folderPath);

            for (int i = 1; i <= 2; i++)
            {
                string docPath = Path.Combine(folderPath, $"Doc{docCounter}.docx");
                CreateSampleDocx(docPath, $"This is the content of document {docCounter} located in {folder}.");
                docCounter++;
            }
        }

        // Prepare the master document that will receive all appended documents.
        Document masterDoc = new Document();

        // Find all DOCX files in the subfolders (recursive search).
        string[] sourceFiles = Directory.GetFiles(baseDir, "*.docx", SearchOption.AllDirectories);
        if (sourceFiles.Length == 0)
            throw new InvalidOperationException("No source DOCX files were found.");

        // Append each source document using UseDestinationStyles.
        foreach (string file in sourceFiles)
        {
            Document srcDoc = new Document(file);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
        }

        // Validate that the master document contains the expected number of sections.
        // Each source document has one section; the master started with one empty section.
        int expectedSections = 1 + sourceFiles.Length;
        if (masterDoc.Sections.Count != expectedSections)
            throw new InvalidOperationException("The merged document does not contain the expected number of sections.");

        // Save the merged document as PDF.
        string outputPdf = Path.Combine(baseDir, "MergedOutput.pdf");
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPdf))
            throw new FileNotFoundException("The PDF output file was not created.", outputPdf);
    }

    // Helper method to create a simple DOCX file with specified text.
    private static void CreateSampleDocx(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath, SaveFormat.Docx);
    }
}
