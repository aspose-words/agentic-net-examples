using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Path to the DOCM (macro‑enabled) file.
        string filePath = "Macro.docm";

        // Detect macro presence without fully loading the document.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"Has macros (detect): {formatInfo.HasMacros}");

        // Load the document.
        Document doc = new Document(filePath);
        Console.WriteLine($"Has macros (document): {doc.HasMacros}");

        // If macros exist, enumerate each VBA module and output its source code.
        if (doc.HasMacros && doc.VbaProject != null)
        {
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                Console.WriteLine($"Module name: {module.Name}");
                Console.WriteLine("Source code:");
                Console.WriteLine(module.SourceCode);
                Console.WriteLine(new string('-', 40));
            }
        }
        else
        {
            Console.WriteLine("No macros found in the document.");
        }
    }
}
