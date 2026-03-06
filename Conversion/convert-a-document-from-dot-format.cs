using System;
using Aspose.Words;

namespace AsposeWordsExamples
{
    class DotConversion
    {
        static void Main()
        {
            // Path to the folder that contains the source .dot file.
            string inputPath = @"C:\Docs\Template.dot";

            // Load the DOT (Word template) document.
            Document doc = new Document(inputPath);

            // Path for the converted document.
            string outputPath = @"C:\Docs\Converted.docx";

            // Save the document in DOCX format.
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
