using System;
using Aspose.Words;

namespace AsposeWordsConversionExample
{
    class Program
    {
        static void Main()
        {
            // Define input and output locations.
            // Replace these with actual paths in your environment.
            string inputPath = @"C:\MyDir\Document.docx";
            string outputPath = @"C:\ArtifactsDir\Document.ConvertToDoc.doc";

            // Load the source document (DOCX) using the Document constructor.
            Document doc = new Document(inputPath);

            // Save the document in the legacy DOC format.
            doc.Save(outputPath, SaveFormat.Doc);
        }
    }
}
