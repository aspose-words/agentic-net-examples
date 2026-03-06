using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDocConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input file (any supported format, e.g., DOCX)
            string inputPath = @"C:\Input\SampleDocument.docx";

            // Output file in legacy DOC format
            string outputPath = @"C:\Output\SampleDocument.doc";

            // Load the source document from the file system
            Document doc = new Document(inputPath);

            // Option 1: Save directly using the SaveFormat enumeration
            doc.Save(outputPath, SaveFormat.Doc);

            // Option 2: Save using DocSaveOptions (uncomment if you need additional options)
            // DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
            // saveOptions.Password = "MyPassword"; // Example: set a password
            // doc.Save(outputPath, saveOptions);
        }
    }
}
