using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the temporary macro‑enabled document.
        const string filePath = "SampleWithMacros.docm";

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Add a VBA module if none exist.
        VbaModuleCollection modules = doc.VbaProject.Modules;
        if (modules.Count == 0)
        {
            // Simple macro source code.
            const string macroCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";

            // Create a module using the parameterless constructor and set its properties.
            VbaModule module = new VbaModule();
            module.Name = "SampleModule";
            module.SourceCode = macroCode;

            modules.Add(module);
        }

        // Save the document in a macro‑enabled format.
        doc.Save(filePath, SaveFormat.Docm);

        // Load the document back.
        Document loadedDoc = new Document(filePath);

        // Access the VBA project.
        VbaProject vbaProject = loadedDoc.VbaProject;
        if (vbaProject != null)
        {
            VbaModuleCollection loadedModules = vbaProject.Modules;

            // Enumerate all modules and output their names and source code.
            foreach (VbaModule mod in loadedModules)
            {
                string source = mod.SourceCode ?? string.Empty;
                Console.WriteLine($"Module Name: {mod.Name}");
                Console.WriteLine("Source Code:");
                Console.WriteLine(source);
                Console.WriteLine(new string('-', 30));
            }
        }
        else
        {
            Console.WriteLine("No VBA project found in the document.");
        }
    }
}
