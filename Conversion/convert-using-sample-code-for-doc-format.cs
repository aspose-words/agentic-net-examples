using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSample
{
    class ConvertToDoc
    {
        static void Main()
        {
            // Define input and output paths (replace with actual paths as needed).
            string inputPath = @"C:\Docs\Document.docx";
            string outputPath = @"C:\Docs\Converted\Document.ConvertToDoc.doc";

            // Load the source document (DOCX, PDF, etc.) using the Document constructor.
            Document doc = new Document(inputPath);

            // Create DocSaveOptions specifying the target format as DOC.
            DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

            // Optional: set a password for the saved DOC file.
            // saveOptions.Password = "MyPassword";

            // Save the document in the older DOC format using the save options.
            doc.Save(outputPath, saveOptions);
        }
    }
}
