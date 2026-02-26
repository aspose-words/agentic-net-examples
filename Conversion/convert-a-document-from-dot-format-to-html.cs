using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOT (Word template) file path
        string inputPath = @"C:\Docs\Template.dot";

        // Output HTML file path
        string outputPath = @"C:\Docs\Template.html";

        // Load the DOT document
        Document doc = new Document(inputPath);

        // Save the document as HTML
        doc.Save(outputPath, SaveFormat.Html);
    }
}
