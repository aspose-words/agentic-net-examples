using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source document (any supported format)
        string sourcePath = @"C:\Docs\sourceFile.doc";

        // Path where the converted DOCX will be saved
        string outputPath = @"C:\Docs\convertedFile.docx";

        // LoadOptions with LoadFormat set to Docx forces the loader to treat the input as a DOCX file.
        // This is useful when the format is known beforehand or when you want to ensure DOCX handling.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Docx
        };

        // Load the document using the specified options.
        Document doc = new Document(sourcePath, loadOptions);

        // Save the loaded document as a DOCX file.
        doc.Save(outputPath);
    }
}
