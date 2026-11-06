using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Add a couple of VBA modules with simple code.
        VbaModuleCollection modules = doc.VbaProject.Modules;

        // Create first module and set its name and source code.
        VbaModule moduleOne = new VbaModule();
        moduleOne.Name = "ModuleOne";
        moduleOne.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from ModuleOne\"\nEnd Sub";
        modules.Add(moduleOne);

        // Create second module and set its name and source code.
        VbaModule moduleTwo = new VbaModule();
        moduleTwo.Name = "ModuleTwo";
        moduleTwo.SourceCode = "Sub ShowInfo()\n    MsgBox \"Info from ModuleTwo\"\nEnd Sub";
        modules.Add(moduleTwo);

        // Save the document to a macro‑enabled format in a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            doc.Save(stream, SaveFormat.Docm);
            stream.Position = 0; // Reset for reading.

            // Load the document back from the stream.
            Document loadedDoc = new Document(stream);

            // Enumerate all VBA modules in the loaded document.
            VbaProject vbaProject = loadedDoc.VbaProject;
            if (vbaProject != null)
            {
                VbaModuleCollection loadedModules = vbaProject.Modules;
                foreach (VbaModule module in loadedModules)
                {
                    // Output module name and its source code line count.
                    string source = module.SourceCode ?? string.Empty;
                    int lineCount = source.Split('\n').Length;
                    Console.WriteLine($"Module: {module.Name}, Lines of code: {lineCount}");
                }
            }
        }
    }
}
