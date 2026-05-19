using System;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace BatchAppendExample
{
    public class Program
    {
        public static void Main()
        {
            // Base directory for sample documents.
            string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
            Directory.CreateDirectory(baseDir);

            // Create subfolders and sample DOCX files.
            string[] subFolders = { "FolderA", "FolderB" };
            int docIndex = 1;
            foreach (string folder in subFolders)
            {
                string folderPath = Path.Combine(baseDir, folder);
                Directory.CreateDirectory(folderPath);

                // Create two documents per subfolder.
                for (int i = 1; i <= 2; i++)
                {
                    string docPath = Path.Combine(folderPath, $"Doc{docIndex}.docx");
                    Document doc = new Document();
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    builder.Writeln($"This is document {docIndex} located in {folder}.");
                    doc.Save(docPath, SaveFormat.Docx);
                    docIndex++;
                }
            }

            // Master document that will receive all appended documents.
            Document masterDoc = new Document();

            // Find all DOCX files in all subfolders.
            string[] sourceFiles = Directory.GetFiles(baseDir, "*.docx", SearchOption.AllDirectories);
            foreach (string file in sourceFiles)
            {
                Document srcDoc = new Document(file);
                masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
            }

            // Validate that the expected number of sections were added.
            // Initial master document has one section; each appended document adds its sections.
            int expectedSections = 1 + sourceFiles.Length;
            if (masterDoc.Sections.Count != expectedSections)
                throw new InvalidOperationException("The merged document does not contain the expected number of sections.");

            // Export the merged document to PDF.
            string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");
            masterDoc.Save(outputPdf, SaveFormat.Pdf);

            // Verify that the PDF file was created.
            if (!File.Exists(outputPdf))
                throw new FileNotFoundException("Failed to create the merged PDF file.", outputPdf);
        }
    }
}
