using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file.
        string filePath = "MyDir\\Document.docx";

        // Load the document (uses the Document constructor – the load rule).
        Document doc = new Document(filePath);

        // Check if the document contains VBA macros using the Document.HasMacros property.
        bool hasMacros = doc.HasMacros;
        Console.WriteLine($"Document.HasMacros: {hasMacros}");

        // Detect macros without fully loading the document using FileFormatUtil.
        FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"FileFormatInfo.HasMacros: {info.HasMacros}");

        // If macros are present, enumerate them via the VBA project.
        if (hasMacros && doc.VbaProject != null)
        {
            foreach (var module in doc.VbaProject.Modules)
            {
                Console.WriteLine($"Module Name: {module.Name}");
                Console.WriteLine("Source Code:");
                Console.WriteLine(module.SourceCode);
                Console.WriteLine(new string('-', 40));
            }
        }
    }
}
