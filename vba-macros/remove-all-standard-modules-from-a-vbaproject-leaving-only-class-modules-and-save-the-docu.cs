using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Add a procedural (standard) module.
        VbaModule procModule = new VbaModule
        {
            Name = "StandardModule1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Test()\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(procModule);

        // Add a document module (also a standard module).
        VbaModule docModule = new VbaModule
        {
            Name = "DocModule",
            Type = VbaModuleType.DocumentModule,
            SourceCode = "Sub DocMacro()\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(docModule);

        // Add a class module (the type we want to keep).
        VbaModule classModule = new VbaModule
        {
            Name = "MyClass",
            Type = VbaModuleType.ClassModule,
            SourceCode = "Public Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(classModule);

        // Remove all modules that are not class modules.
        VbaModuleCollection modules = doc.VbaProject.Modules;
        for (int i = modules.Count - 1; i >= 0; i--)
        {
            VbaModule module = modules[i];
            if (module.Type != VbaModuleType.ClassModule)
                modules.Remove(module);
        }

        // Save the document in a macro‑enabled format.
        doc.Save("Output.docm");
    }
}
