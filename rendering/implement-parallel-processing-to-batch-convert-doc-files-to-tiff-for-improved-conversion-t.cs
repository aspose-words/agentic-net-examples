using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeBatchConvert");
        if (Directory.Exists(workFolder))
            Directory.Delete(workFolder, true);
        Directory.CreateDirectory(workFolder);

        // Number of sample DOCX files to generate.
        int sampleCount = 5;

        // Generate sample DOCX documents.
        for (int i = 1; i <= sampleCount; i++)
        {
            // Create a new blank document.
            Document doc = new Document();
            // Add a paragraph with sample text.
            var builder = new Aspose.Words.DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}.");
            // Save the document as DOCX.
            string docPath = Path.Combine(workFolder, $"Sample{i}.docx");
            doc.Save(docPath);
        }

        // Get all DOCX files in the working folder.
        string[] docFiles = Directory.GetFiles(workFolder, "*.docx");

        // Convert each DOCX to a multipage TIFF in parallel.
        Parallel.ForEach(docFiles, docFile =>
        {
            // Load the source document.
            Document srcDoc = new Document(docFile);

            // Determine output TIFF path.
            string tiffPath = Path.ChangeExtension(docFile, ".tiff");

            // Save the document as TIFF (each page becomes a frame).
            srcDoc.Save(tiffPath, SaveFormat.Tiff);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF for '{docFile}'.");
        });

        // Optional: Verify that the number of TIFF files matches the number of DOCX files.
        int tiffCount = Directory.GetFiles(workFolder, "*.tiff").Length;
        if (tiffCount != sampleCount)
            throw new InvalidOperationException("Mismatch between source DOCX files and generated TIFF files.");

        // Cleanup: delete the temporary folder (optional).
        // Directory.Delete(workFolder, true);
    }
}
