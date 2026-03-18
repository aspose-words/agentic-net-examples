using System;
using Aspose.Words;

class ReplaceTextInDocx
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\ModifiedDocument.docx";

        // Load the existing document from the file system.
        Document doc = new Document(inputPath);

        // The literal text to find and its replacement.
        string textToFind = "_FullName_";
        string replacementText = "John Doe";

        // Perform a case‑insensitive, whole‑document replace.
        // This uses the Range.Replace(string pattern, string replacement) method.
        doc.Range.Replace(textToFind, replacementText);

        // Save the updated document.
        doc.Save(outputPath);
    }
}
