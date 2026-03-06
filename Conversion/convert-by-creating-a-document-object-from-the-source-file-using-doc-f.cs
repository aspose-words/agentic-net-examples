using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    public class DocConverter
    {
        /// <summary>
        /// Loads a document from the specified file and saves it in the legacy DOC format.
        /// </summary>
        /// <param name="sourcePath">Full path to the source document (any format supported by Aspose.Words).</param>
        /// <param name="destPath">Full path where the DOC file will be saved.</param>
        public void ConvertToDoc(string sourcePath, string destPath)
        {
            // Load the source document. The constructor automatically detects the file format.
            Document doc = new Document(sourcePath);

            // Save the document in the Microsoft Word 97‑2007 DOC format.
            doc.Save(destPath, SaveFormat.Doc);
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = @"C:\Input\sample.pdf";   // replace with your source file
            string destinationFile = @"C:\Output\sample.doc";

            DocConverter converter = new DocConverter();
            converter.ConvertToDoc(sourceFile, destinationFile);

            Console.WriteLine("Conversion completed.");
        }
    }
}
