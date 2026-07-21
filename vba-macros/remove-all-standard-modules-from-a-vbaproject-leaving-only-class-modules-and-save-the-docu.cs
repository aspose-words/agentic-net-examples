using System;
using System.Collections.Generic;
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
            VbaProject vbaProject = new VbaProject
            {
                Name = "SampleProject"
            };
            doc.VbaProject = vbaProject;

            // Add a procedural module.
            VbaModule procModule = new VbaModule
            {
                Name = "ProceduralModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Sub ProcMacro()\n    MsgBox \"Procedural\"\nEnd Sub"
            };
            doc.VbaProject.Modules.Add(procModule);

            // Add a document module.
            VbaModule docModule = new VbaModule
            {
                Name = "DocumentModule",
                Type = VbaModuleType.DocumentModule,
                SourceCode = "Sub DocMacro()\n    MsgBox \"Document\"\nEnd Sub"
            };
            doc.VbaProject.Modules.Add(docModule);

            // Add a designer module.
            VbaModule designerModule = new VbaModule
            {
                Name = "DesignerModule",
                Type = VbaModuleType.DesignerModule,
                SourceCode = "Sub DesignerMacro()\n    MsgBox \"Designer\"\nEnd Sub"
            };
            doc.VbaProject.Modules.Add(designerModule);

            // Add a class module (this one should be kept).
            VbaModule classModule = new VbaModule
            {
                Name = "ClassModule",
                Type = VbaModuleType.ClassModule,
                SourceCode = "Public Sub ClassMacro()\n    MsgBox \"Class\"\nEnd Sub"
            };
            doc.VbaProject.Modules.Add(classModule);

            // Identify and remove all modules that are not class modules.
            List<VbaModule> modulesToRemove = new List<VbaModule>();
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                if (module.Type != VbaModuleType.ClassModule)
                {
                    modulesToRemove.Add(module);
                }
            }

            foreach (VbaModule module in modulesToRemove)
            {
                doc.VbaProject.Modules.Remove(module);
            }

            // Save the resulting document as a macro-enabled file.
            doc.Save("Result.docm");
        }
    }
}
