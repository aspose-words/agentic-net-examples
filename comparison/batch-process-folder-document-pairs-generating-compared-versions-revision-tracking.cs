using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class BatchDocumentComparer
{
    static void Main()
    {
        // Folder containing the document pairs.
        string inputFolder = @"C:\Docs\Input";
        // Folder where the compared documents will be saved.
        string outputFolder = @"C:\Docs\Output";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Find all original documents (assumed to end with "_original").
        foreach (string originalPath in Directory.GetFiles(inputFolder, "*_original.*"))
        {
            // Derive the base name without the "_original" suffix.
            string fileName = Path.GetFileNameWithoutExtension(originalPath);
            string baseName = fileName.Substring(0, fileName.LastIndexOf("_original", StringComparison.Ordinal));
            string extension = Path.GetExtension(originalPath);

            // Construct the expected edited document path.
            string editedPath = Path.Combine(inputFolder, $"{baseName}_edited{extension}");

            // Skip if the edited counterpart does not exist.
            if (!File.Exists(editedPath))
                continue;

            // Load the original and edited documents.
            Document docOriginal = new Document(originalPath);
            Document docEdited = new Document(editedPath);

            // Documents must not contain revisions before comparison.
            if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
            {
                // Compare the documents, generating revisions in the original document.
                docOriginal.Compare(docEdited, "BatchComparer", DateTime.Now);
            }

            // Save the compared document with revisions to the output folder.
            string outputPath = Path.Combine(outputFolder, $"{baseName}_compared{extension}");
            docOriginal.Save(outputPath);
        }
    }
}
