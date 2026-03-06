using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsHtmlConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input file path – can be any format supported by Aspose.Words (docx, rtf, pdf, etc.).
            string inputFile = @"C:\Input\SampleDocument.docx";

            // Output file path – the extension determines the format, but we also explicitly specify Html.
            string outputFile = @"C:\Output\SampleDocument.html";

            // Load the document from the file system. The constructor automatically detects the format.
            Document doc = new Document(inputFile);

            // Save the document as HTML. The SaveFormat enum value ensures the correct output type.
            doc.Save(outputFile, SaveFormat.Html);
        }
    }
}
