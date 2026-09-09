using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for the sample input documents.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        Directory.CreateDirectory(baseDir);

        // Create two subfolders and seed each with two DOCX files.
        string[] subFolders = { "FolderA", "FolderB" };
        int docIndex = 1;

        foreach (string folder in subFolders)
        {
            string folderPath = Path.Combine(baseDir, folder);
            Directory.CreateDirectory(folderPath);

            for (int i = 1; i <= 2; i++)
            {
                // Create a simple document with identifiable content.
                Document srcDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(srcDoc);
                builder.Writeln($"Document {docIndex} from {folder}");
                string docPath = Path.Combine(folderPath, $"Doc{docIndex}.docx");
                srcDoc.Save(docPath, SaveFormat.Docx);
                docIndex++;
            }
        }

        // Master document that will receive all appended documents.
        Document masterDoc = new Document();

        // Find every DOCX file in all subfolders and append them.
        string[] docFiles = Directory.GetFiles(baseDir, "*.docx", SearchOption.AllDirectories);
        foreach (string filePath in docFiles)
        {
            Document srcDoc = new Document(filePath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
        }

        // Export the merged document to PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF output file was not created.");

        // Validate that the merged document contains the expected text from each source file.
        string mergedText = masterDoc.GetText();

        foreach (string filePath in docFiles)
        {
            // Extract the numeric part from the file name (e.g., "Doc12" -> 12).
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            string numberPart = new string(fileNameWithoutExt.SkipWhile(c => !char.IsDigit(c))
                                                             .TakeWhile(char.IsDigit)
                                                             .ToArray());

            if (!int.TryParse(numberPart, out int number))
                continue; // Skip if we cannot parse the number.

            // Determine the folder name (FolderA or FolderB).
            string folderName = new DirectoryInfo(Path.GetDirectoryName(filePath)!).Name;

            // Build the expected snippet that was written into the source document.
            string expectedSnippet = $"Document {number} from {folderName}";

            if (!mergedText.Contains(expectedSnippet))
                throw new InvalidOperationException($"Merged document is missing content: \"{expectedSnippet}\".");
        }
    }
}
