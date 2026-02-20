using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load a macro‑enabled Word document (DOCM). Adjust the path as needed.
        Document doc = new Document("Input.docm");

        // Verify that the document actually contains VBA macros.
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Access the VBA project associated with the document.
        VbaProject vbaProject = doc.VbaProject;

        // Output basic information about the VBA project.
        Console.WriteLine($"Project name: {vbaProject.Name}");
        Console.WriteLine($"Signed: {vbaProject.IsSigned}");
        Console.WriteLine($"Code page: {vbaProject.CodePage}");
        Console.WriteLine($"Modules count: {vbaProject.Modules.Count}");

        // Iterate through each VBA module and display its source code.
        foreach (VbaModule module in vbaProject.Modules)
        {
            Console.WriteLine($"--- Module: {module.Name} (Type: {module.Type}) ---");
            Console.WriteLine(module.SourceCode);
            Console.WriteLine();
        }
    }
}
