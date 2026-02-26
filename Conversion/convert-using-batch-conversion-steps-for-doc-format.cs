using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchDocConverter
{
    static void Main()
    {
        // Folder containing source documents of various formats.
        string inputFolder = @"C:\Input";

        // Folder where the converted DOC files will be saved.
        string outputFolder = @"C:\Output";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Iterate over each file in the input folder.
        foreach (string sourcePath in Directory.GetFiles(inputFolder))
        {
            // Load the source document (auto-detects format).
            Document doc = new Document(sourcePath);

            // Create save options for the legacy DOC format.
            DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

            // Construct the output file name with a .doc extension.
            string targetPath = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(sourcePath) + ".doc");

            // Save the document as DOC using the specified options.
            doc.Save(targetPath, saveOptions);
        }
    }
}
