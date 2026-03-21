using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class BatchDocumentComparer
{
    static void Main()
    {
        // Use folders relative to the executable location.
        string baseDir = AppContext.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Find all original documents (assumed to end with "_original").
        foreach (string originalPath in Directory.GetFiles(inputFolder, "*_original.*"))
        {
            // Derive the base name without the "_original" suffix.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
            if (!fileNameWithoutExt.EndsWith("_original", StringComparison.OrdinalIgnoreCase))
                continue; // Safety check.

            string baseName = fileNameWithoutExt.Substring(0,
                fileNameWithoutExt.Length - "_original".Length);
            string extension = Path.GetExtension(originalPath);

            // Construct the expected edited document path.
            string editedPath = Path.Combine(inputFolder, $"{baseName}_edited{extension}");

            // Skip if the edited counterpart does not exist.
            if (!File.Exists(editedPath))
                continue;

            // Load the original and edited documents.
            Document docOriginal = new Document(originalPath);
            Document docEdited = new Document(editedPath);

            // Compare the documents, generating revisions in the original document.
            docOriginal.Compare(docEdited, "BatchComparer", DateTime.Now);

            // Save the compared document with revisions to the output folder.
            string outputPath = Path.Combine(outputFolder, $"{baseName}_compared{extension}");
            docOriginal.Save(outputPath);
        }
    }
}
