using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the document (use a .docm or .docx that may contain macros)
        Document doc = new Document("MyDir\\MacroDocument.docm");

        // Determine whether the document has a VBA project (macros)
        if (doc.HasMacros)
        {
            Console.WriteLine("Document contains VBA macros.");

            // Access the VBA project attached to the document
            VbaProject vbaProject = doc.VbaProject;

            // Output basic information about the VBA project
            Console.WriteLine($"Project name: {vbaProject.Name}");
            Console.WriteLine($"Signed: {vbaProject.IsSigned}");
            Console.WriteLine($"Code page: {vbaProject.CodePage}");
            Console.WriteLine($"Modules count: {vbaProject.Modules.Count}");

            // Iterate through each VBA module and display its source code
            foreach (VbaModule module in vbaProject.Modules)
            {
                Console.WriteLine($"--- Module: {module.Name} ---");
                Console.WriteLine(module.SourceCode);
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine("Document does not contain any VBA macros.");
        }
    }
}
