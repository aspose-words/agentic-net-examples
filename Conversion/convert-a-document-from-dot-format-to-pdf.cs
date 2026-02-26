using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOT template file.
            string dotFilePath = @"C:\Input\Template.dot";

            // Path where the resulting PDF will be saved.
            string pdfFilePath = @"C:\Output\Converted.pdf";

            // Load the DOT document. The Document constructor automatically detects the format.
            Document doc = new Document(dotFilePath);

            // Save the document as PDF. The file extension determines the save format.
            doc.Save(pdfFilePath);
        }
    }
}
