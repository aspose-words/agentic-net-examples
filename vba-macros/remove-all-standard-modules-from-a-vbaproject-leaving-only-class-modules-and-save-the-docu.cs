using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // Add a procedural module.
        VbaModule procModule = new VbaModule();
        procModule.Name = "StandardProcedural";
        procModule.Type = VbaModuleType.ProceduralModule;
        procModule.SourceCode = "Sub ProcMacro()\n    MsgBox \"Procedural\"\nEnd Sub";
        doc.VbaProject.Modules.Add(procModule);

        // Add a document module.
        VbaModule docModule = new VbaModule();
        docModule.Name = "StandardDocument";
        docModule.Type = VbaModuleType.DocumentModule;
        docModule.SourceCode = "Sub DocMacro()\n    MsgBox \"Document\"\nEnd Sub";
        doc.VbaProject.Modules.Add(docModule);

        // Add a class module (this one should be kept).
        VbaModule classModule = new VbaModule();
        classModule.Name = "MyClass";
        classModule.Type = VbaModuleType.ClassModule;
        classModule.SourceCode = "Public Sub ClassMacro()\n    MsgBox \"Class\"\nEnd Sub";
        doc.VbaProject.Modules.Add(classModule);

        // Add a designer module.
        VbaModule designerModule = new VbaModule();
        designerModule.Name = "StandardDesigner";
        designerModule.Type = VbaModuleType.DesignerModule;
        designerModule.SourceCode = "Sub DesignerMacro()\n    MsgBox \"Designer\"\nEnd Sub";
        doc.VbaProject.Modules.Add(designerModule);

        // Remove all modules that are not class modules.
        List<VbaModule> modulesToRemove = new List<VbaModule>();
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            if (module.Type != VbaModuleType.ClassModule)
                modulesToRemove.Add(module);
        }

        foreach (VbaModule module in modulesToRemove)
        {
            doc.VbaProject.Modules.Remove(module);
        }

        // Save the resulting document as a macro-enabled file.
        doc.Save("Result.docm");
    }
}
