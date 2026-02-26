using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSample
{
    class Program
    {
        static void Main()
        {
            // Define input and output paths.
            string inputPath = @"C:\Docs\SampleInput.docx";
            string outputPath = @"C:\Docs\ConvertedOutput.doc";

            // Load an existing document (create/load lifecycle).
            Document doc = new Document(inputPath);

            // Create save options for the legacy DOC format.
            DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

            // Optionally set a password or other options here.
            // saveOptions.Password = "MyPassword";

            // Save the document as DOC using the save options (save lifecycle).
            doc.Save(outputPath, saveOptions);
        }
    }
}
