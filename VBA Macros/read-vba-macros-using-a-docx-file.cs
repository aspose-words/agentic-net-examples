using System;
using Aspose.Words;
using Aspose.Words.Vba;

class ReadVbaMacros
{
    static void Main()
    {
        // Path to the Word document (macro-enabled .docm or .docx if it somehow contains macros)
        string filePath = @"C:\Docs\Sample.docm";

        // Load the document using Aspose.Words Document constructor (lifecycle rule)
        Document doc = new Document(filePath);

        // Check whether the document contains a VBA project (macros)
        if (doc.HasMacros)
        {
            // Access the VBA project attached to the document
            VbaProject vbaProject = doc.VbaProject;

            // Iterate through each VBA module and output its source code
            foreach (VbaModule module in vbaProject.Modules)
            {
                Console.WriteLine($"--- Module: {module.Name} ---");
                Console.WriteLine(module.SourceCode);
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine("The document does not contain any VBA macros.");
        }
    }
}
