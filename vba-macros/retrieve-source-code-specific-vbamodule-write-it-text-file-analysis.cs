using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the macro‑enabled document.
        Document doc = new Document("Input.docm");

        // Verify that the document actually contains a VBA project.
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Get the VBA project from the document.
        VbaProject vbaProject = doc.VbaProject;

        // Choose the module whose source code you want to extract.
        // You can use the module name or its zero‑based index.
        string targetModuleName = "Module1";               // change as needed
        VbaModule module = vbaProject.Modules[targetModuleName];

        if (module == null)
        {
            Console.WriteLine($"Module '{targetModuleName}' was not found in the VBA project.");
            return;
        }

        // Retrieve the VBA source code from the selected module.
        string sourceCode = module.SourceCode;

        // Write the source code to a plain‑text file for analysis.
        File.WriteAllText("ModuleSource.txt", sourceCode);

        Console.WriteLine($"Source code of module '{targetModuleName}' has been saved to ModuleSource.txt");
    }
}
