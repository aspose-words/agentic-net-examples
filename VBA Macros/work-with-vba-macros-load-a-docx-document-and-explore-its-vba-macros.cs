using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the DOCX/DOCM file. Replace the path with your actual file location.
        Document doc = new Document("Macro.docm");

        // Check whether the document contains any VBA macros.
        Console.WriteLine($"Has macros: {doc.HasMacros}");

        // If macros are present, explore the VBA project.
        if (doc.HasMacros && doc.VbaProject != null)
        {
            VbaProject vbaProject = doc.VbaProject;

            // Basic information about the VBA project.
            Console.WriteLine($"VBA Project Name: {vbaProject.Name}");
            Console.WriteLine($"Is Signed: {vbaProject.IsSigned}");
            Console.WriteLine($"Modules count: {vbaProject.Modules.Count}");

            // Iterate through each VBA module and display its name and source code.
            foreach (VbaModule module in vbaProject.Modules)
            {
                Console.WriteLine($"Module Name: {module.Name}");
                Console.WriteLine("Source Code:");
                Console.WriteLine(module.SourceCode);
                Console.WriteLine(new string('-', 40));
            }
        }
        else
        {
            Console.WriteLine("No VBA project found in the document.");
        }
    }
}
