using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file that contains the placeholder.
        const string inputPath = @"C:\Docs\Template.docx";

        // Path where the resulting document will be saved.
        const string outputPath = @"C:\Docs\Result.docx";

        // The placeholder text to be replaced (e.g., _FullName_).
        const string placeholder = "_FullName_";

        // The dynamic content that will replace the placeholder.
        const string replacement = "John Doe";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Perform a find-and-replace operation on the whole document range.
        // This uses the Range.Replace(string, string) method as defined in the Aspose.Words API.
        doc.Range.Replace(placeholder, replacement);

        // Save the modified document.
        doc.Save(outputPath);
    }
}
