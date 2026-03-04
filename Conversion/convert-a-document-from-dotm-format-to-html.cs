using System;
using Aspose.Words;
using Aspose.Words.Saving;

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

            // Save the document in HTML format using the Save method that accepts a file name
            // and a SaveFormat enumeration value.
            doc.Save(outputFile, SaveFormat.Html);
        }
    }
}
