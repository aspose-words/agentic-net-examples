using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace RemoveStandardModulesExample
{
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

            // Add a procedural module (standard module).
            VbaModule proceduralModule = new VbaModule
            {
                Name = "StandardModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Sub ProcMacro()\n    MsgBox \"Procedural\"\nEnd Sub"
            };
            doc.VbaProject.Modules.Add(proceduralModule);

            // Add a document module (standard module).
            VbaModule documentModule = new VbaModule
            {
                Name = "DocModule",
                Type = VbaModuleType.DocumentModule,
                SourceCode = "Sub DocMacro()\n    MsgBox \"Document\"\nEnd Sub"
            };
            doc.VbaProject.Modules.Add(documentModule);

            // Add a class module (the one we want to keep).
            VbaModule classModule = new VbaModule
            {
                Name = "MyClass",
                Type = VbaModuleType.ClassModule,
                SourceCode = "Public Sub ClassMacro()\n    MsgBox \"Class\"\nEnd Sub"
            };
            doc.VbaProject.Modules.Add(classModule);

            // Remove all modules that are not class modules.
            VbaModuleCollection modules = doc.VbaProject.Modules;
            for (int i = modules.Count - 1; i >= 0; i--)
            {
                VbaModule module = modules[i];
                if (module.Type != VbaModuleType.ClassModule)
                {
                    modules.Remove(module);
                }
            }

            // Save the document in a macro‑enabled format.
            doc.Save("Result.docm");
        }
    }
}
