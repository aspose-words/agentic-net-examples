using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourcePath = @"C:\Docs\Document.docx";

            // Path where the converted DOC file will be saved.
            string outputPath = @"C:\Output\Document.Converted.doc";

            // Load the existing DOCX document.
            Document doc = new Document(sourcePath);

            // Save the document in the legacy DOC format.
            doc.Save(outputPath, SaveFormat.Doc);
        }
    }
}
