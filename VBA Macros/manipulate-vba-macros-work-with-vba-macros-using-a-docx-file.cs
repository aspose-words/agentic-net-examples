using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Path to the input document (must be a macro‑enabled format, e.g., .docm).
        string inputPath = @"C:\Docs\InputDocument.docm";
        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\OutputDocument.docm";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Check whether the document already contains a VBA project (macros).
        if (doc.HasMacros)
        {
            Console.WriteLine("Document already contains VBA macros.");
        }
        else
        {
            Console.WriteLine("Document does not contain VBA macros. Creating a new VBA project.");
            // Create a new VBA project and assign it to the document.
            VbaProject project = new VbaProject
            {
                Name = "AsposeDemoProject",
                CodePage = 1252 // Typical Windows Latin code page.
            };
            doc.VbaProject = project;
        }

        // Access the VBA project (it is guaranteed to exist now).
        VbaProject vbaProject = doc.VbaProject;

        // Create a new procedural module.
        VbaModule newModule = new VbaModule
        {
            Name = "DemoModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words VBA!""
End Sub"
        };

        // Add the module to the project's module collection.
        vbaProject.Modules.Add(newModule);
        Console.WriteLine($"Added module '{newModule.Name}' to the VBA project.");

        // Optionally, modify an existing module if one is present.
        if (vbaProject.Modules.Count > 0)
        {
            // Example: prepend a comment to the first module's source code.
            VbaModule firstModule = vbaProject.Modules[0];
            firstModule.SourceCode = "' Updated by Aspose.Words\n" + firstModule.SourceCode;
            Console.WriteLine($"Updated source code of module '{firstModule.Name}'.");
        }

        // Save the document in a macro‑enabled format.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
