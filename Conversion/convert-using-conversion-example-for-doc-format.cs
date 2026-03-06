using System;
using Aspose.Words;

namespace AsposeWordsConversionExample
{
    class Program
    {
        static void Main()
        {
            // Define input and output paths.
            string inputPath = @"MyDir\Document.docx";          // Source document (any supported format)
            string outputPath = @"ArtifactsDir\Document.ConvertToDoc.doc"; // Destination in DOC format

            // Load the source document.
            Document doc = new Document(inputPath);

            // Save the document in the legacy Microsoft Word 97‑2007 DOC format.
            doc.Save(outputPath, SaveFormat.Doc);
        }
    }
}
