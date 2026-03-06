using System;
using Aspose.Words;

namespace DotmToHtmlConverter
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOTM (macro‑enabled template) file.
            string sourcePath = @"C:\Docs\Template.dotm";

            // Path where the resulting HTML file will be saved.
            string targetPath = @"C:\Docs\Template.html";

            // Load the DOTM document. The Document constructor automatically detects the format.
            Document doc = new Document(sourcePath);

            // Save the document as HTML using the explicit SaveFormat enumeration.
            doc.Save(targetPath, SaveFormat.Html);
        }
    }
}
