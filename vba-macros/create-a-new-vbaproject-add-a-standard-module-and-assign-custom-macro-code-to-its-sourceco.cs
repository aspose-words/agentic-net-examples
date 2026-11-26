using System;
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

        // Define custom macro code.
        string macroCode = "Sub MyMacro()\n    MsgBox \"Hello from macro\"\nEnd Sub";

        // Create a standard module and assign its name and source code.
        VbaModule module = new VbaModule
        {
            Name = "StandardModule1",
            SourceCode = macroCode
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        string outputPath = "MacroDocument.docm";
        doc.Save(outputPath, SaveFormat.Docm);

        // Validate that the module was added correctly.
        VbaModule addedModule = doc.VbaProject.Modules["StandardModule1"];
        if (addedModule != null && addedModule.SourceCode == macroCode)
        {
            Console.WriteLine("Macro module added successfully.");
        }
        else
        {
            Console.WriteLine("Failed to add macro module.");
        }
    }
}
