using System;
using Aspose.Words;

namespace AsposeWordsExamples
{
    class ConvertToPdf
    {
        static void Main()
        {
            // Define the folder where the source document is located.
            // Adjust the path as needed for your environment.
            string MyDir = @"C:\Docs\Input\";

            // Define the folder where the resulting PDF will be saved.
            // Adjust the path as needed for your environment.
            string ArtifactsDir = @"C:\Docs\Output\";

            // Load an existing DOCX document from the file system.
            // This uses the Document(string fileName) constructor rule.
            Document doc = new Document(MyDir + "Document.docx");

            // Save the document as PDF.
            // The Save(string fileName) method automatically determines the format from the extension.
            doc.Save(ArtifactsDir + "Document.ConvertToPdf.pdf");
        }
    }
}
