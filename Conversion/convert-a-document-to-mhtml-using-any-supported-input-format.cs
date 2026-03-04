using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMhtmlConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input file path – can be any format supported by Aspose.Words (e.g., .docx, .pdf, .rtf, etc.).
            string inputFile = @"C:\Input\SampleDocument.docx";

            // Output file path – the extension determines the format, but we will also specify MHTML explicitly.
            string outputFile = @"C:\Output\SampleDocument.mht";

            // Load the source document. The constructor automatically detects the format.
            Document doc = new Document(inputFile);

            // Create HtmlSaveOptions with the SaveFormat set to Mhtml.
            // This tells Aspose.Words to save the document as a web archive (MHTML).
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);

            // Save the document to the specified output path using the MHTML options.
            doc.Save(outputFile, saveOptions);

            Console.WriteLine("Document successfully converted to MHTML.");
        }
    }
}
