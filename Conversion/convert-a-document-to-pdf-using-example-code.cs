using System;
using Aspose.Words;

namespace AsposeWordsExamples
{
    class ConvertToPdf
    {
        static void Main()
        {
            // Define the folder where the source document is located.
            // Replace with the actual path to your input .docx file.
            string MyDir = @"C:\InputDocuments\";

            // Define the folder where the converted PDF will be saved.
            // Replace with the desired output folder path.
            string ArtifactsDir = @"C:\OutputDocuments\";

            // Load an existing Word document from the file system.
            // This uses the Document(string fileName) constructor.
            Document doc = new Document(MyDir + "Document.docx");

            // Save the document as PDF.
            // The Save(string fileName) overload automatically determines the format from the extension.
            doc.Save(ArtifactsDir + "Document.ConvertToPdf.pdf");
        }
    }
}
