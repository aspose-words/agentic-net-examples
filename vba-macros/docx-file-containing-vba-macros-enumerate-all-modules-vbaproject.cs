using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the macro-enabled document.
        Document doc = new Document("MyDir/VBA project.docm");

        // Verify that the document actually contains VBA macros.
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Access the VBA project inside the document.
        VbaProject vbaProject = doc.VbaProject;

        // Get the collection of VBA modules.
        VbaModuleCollection modules = vbaProject.Modules;

        Console.WriteLine($"VBA Project Name: {vbaProject.Name}");
        Console.WriteLine($"Number of modules: {modules.Count}");

        // Enumerate each module and output its details.
        foreach (VbaModule module in modules)
        {
            Console.WriteLine($"Module Name: {module.Name}");
            Console.WriteLine($"Module Type: {module.Type}");
            Console.WriteLine("Source Code:");
            Console.WriteLine(module.SourceCode);
            Console.WriteLine(new string('-', 40));
        }
    }
}
