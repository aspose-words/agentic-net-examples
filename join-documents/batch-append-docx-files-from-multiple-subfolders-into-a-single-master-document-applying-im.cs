using System;
using System.IO;
using Aspose.Words;

namespace BatchAppendExample
{
    public class Program
    {
        public static void Main()
        {
            // Root folder for sample documents
            string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");

            // Clean previous run data
            if (Directory.Exists(rootFolder))
                Directory.Delete(rootFolder, true);
            Directory.CreateDirectory(rootFolder);

            // Create subfolders with sample DOCX files
            CreateSampleDocs(rootFolder, "FolderA", 2);
            CreateSampleDocs(rootFolder, "FolderB", 2);

            // Master document that will receive all appended content
            Document masterDoc = new Document();

            // Append every DOCX file from each subfolder using destination styles
            foreach (string subFolder in Directory.GetDirectories(rootFolder))
            {
                foreach (string filePath in Directory.GetFiles(subFolder, "*.docx"))
                {
                    Document srcDoc = new Document(filePath);
                    masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
                }
            }

            // Save the merged document as PDF
            string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");
            masterDoc.Save(outputPdf, SaveFormat.Pdf);

            // Validate PDF creation
            if (!File.Exists(outputPdf))
                throw new InvalidOperationException("Merged PDF was not created.");

            // Validate that the merged document contains sections from all source files
            int expectedSections = CountSourceDocs(rootFolder);
            if (masterDoc.Sections.Count < expectedSections)
                throw new InvalidOperationException("Merged document does not contain all source sections.");

            // Optional: also save as DOCX for manual inspection
            string outputDocx = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.docx");
            masterDoc.Save(outputDocx, SaveFormat.Docx);
        }

        private static void CreateSampleDocs(string rootFolder, string subFolderName, int fileCount)
        {
            string subFolderPath = Path.Combine(rootFolder, subFolderName);
            Directory.CreateDirectory(subFolderPath);

            for (int i = 1; i <= fileCount; i++)
            {
                string filePath = Path.Combine(subFolderPath, $"Sample_{subFolderName}_{i}.docx");
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"This is sample document {i} in {subFolderName}.");
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln($"Heading in {subFolderName} document {i}");
                doc.Save(filePath, SaveFormat.Docx);
            }
        }

        private static int CountSourceDocs(string rootFolder)
        {
            int count = 0;
            foreach (string subFolder in Directory.GetDirectories(rootFolder))
            {
                count += Directory.GetFiles(subFolder, "*.docx").Length;
            }
            return count;
        }
    }
}
