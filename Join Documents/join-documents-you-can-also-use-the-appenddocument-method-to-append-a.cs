using System;
using Aspose.Words;

namespace AsposeWordsJoinExample
{
    class Program
    {
        static void Main()
        {
            // Path to the destination (output) document.
            string outputPath = "JoinedDocument.docx";

            // Paths to the source documents that will be joined.
            string sourcePath1 = "Source1.docx";
            string sourcePath2 = "Source2.docx";

            // Create a new blank document that will serve as the destination.
            Document dstDoc = new Document();

            // Load the first source document from file.
            Document srcDoc1 = new Document(sourcePath1);
            // Append the first source document to the destination.
            // KeepSourceFormatting preserves the original formatting of the source.
            dstDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);

            // Load the second source document from file.
            Document srcDoc2 = new Document(sourcePath2);
            // Append the second source document to the destination.
            dstDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document to the specified output file.
            dstDoc.Save(outputPath);
        }
    }
}
