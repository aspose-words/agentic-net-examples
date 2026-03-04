using System;
using Aspose.Words;

namespace DotToPdfConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOT (Word template) file.
            string inputPath = @"C:\Docs\Template.dot";

            // Path where the resulting PDF will be saved.
            string outputPath = @"C:\Docs\Converted.pdf";

            // Load the DOT document. The constructor automatically detects the format.
            Document doc = new Document(inputPath);

            // Save the document as PDF. The SaveFormat enum explicitly specifies PDF output.
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
