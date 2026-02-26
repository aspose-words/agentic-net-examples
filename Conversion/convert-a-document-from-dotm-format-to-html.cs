using System;
using Aspose.Words;

namespace DotmToHtmlConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOTM (macro‑enabled template) file.
            string inputFile = @"C:\Docs\Template.dotm";

            // Path where the resulting HTML file will be saved.
            string outputFile = @"C:\Docs\Template.html";

            // Load the DOTM document using the Document(string) constructor.
            Document doc = new Document(inputFile);

            // Save the loaded document in HTML format.
            // The Save method with (string, SaveFormat) matches the provided rule.
            doc.Save(outputFile, SaveFormat.Html);
        }
    }
}
