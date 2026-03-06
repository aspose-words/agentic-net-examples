using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOC file.
            string sourcePath = @"C:\Docs\SourceDocument.doc";

            // Specify that the document to be loaded is in DOC format.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Doc, string.Empty, string.Empty);

            // Load the document using the constructor that accepts a file name and LoadOptions.
            Document doc = new Document(sourcePath, loadOptions);

            // Example: save the loaded document as DOCX (optional, demonstrates saving).
            string outputPath = @"C:\Docs\ConvertedDocument.docx";
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
