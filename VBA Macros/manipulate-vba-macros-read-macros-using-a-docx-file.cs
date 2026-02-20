using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the DOCM (macro-enabled) document.
        // Replace the path with the actual location of your file.
        Document doc = new Document("InputDocument.docm");

        // Verify that the document contains a VBA project (macros).
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Access the VBA project.
        VbaProject vbaProject = doc.VbaProject;

        // Output basic information about the VBA project.
        Console.WriteLine($"Project Name: {vbaProject.Name}");
        Console.WriteLine($"Signed: {vbaProject.IsSigned}");
        Console.WriteLine($"Code Page: {vbaProject.CodePage}");
        Console.WriteLine($"Modules Count: {vbaProject.Modules.Count()}");
        Console.WriteLine();

        // Iterate through each VBA module and display its name and source code.
        foreach (VbaModule module in vbaProject.Modules)
        {
            Console.WriteLine($"Module Name: {module.Name}");
            Console.WriteLine("Source Code:");
            Console.WriteLine(module.SourceCode);
            Console.WriteLine(new string('-', 40));
        }
    }
}
