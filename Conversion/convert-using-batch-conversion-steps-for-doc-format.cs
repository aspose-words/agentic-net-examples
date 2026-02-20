using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchDocConverter
{
    static void Main()
    {
        // Folder containing source documents (any supported format)
        string inputFolder = @"C:\Input";

        // Folder where converted DOC files will be saved
        string outputFolder = @"C:\Output";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each file in the input folder
        foreach (string inputFilePath in Directory.GetFiles(inputFolder))
        {
            // Load the document; format is detected automatically
            Document document = new Document(inputFilePath);

            // Configure save options for the legacy DOC format
            DocSaveOptions saveOptions = new DocSaveOptions
            {
                SaveFormat = SaveFormat.Doc
            };

            // Construct the output file name with .doc extension
            string outputFileName = Path.GetFileNameWithoutExtension(inputFilePath) + ".doc";
            string outputFilePath = Path.Combine(outputFolder, outputFileName);

            // Save the document as DOC
            document.Save(outputFilePath, saveOptions);
        }
    }
}
