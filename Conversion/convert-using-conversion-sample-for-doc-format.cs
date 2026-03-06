using System;
using Aspose.Words;

namespace AsposeWordsConversionSample
{
    class Program
    {
        static void Main()
        {
            // Define input and output paths.
            string inputPath = @"MyDir\Document.docx";          // Source document (any supported format)
            string outputPath = @"ArtifactsDir\Document.ConvertToDoc.doc"; // Target DOC format file

            // Load the source document.
            Document doc = new Document(inputPath);

            // Save the document in the legacy DOC format.
            doc.Save(outputPath, SaveFormat.Doc);
        }
    }
}
