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

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // Add a procedural module.
        VbaModule procModule = new VbaModule
        {
            Name = "StandardModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from procedural module\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(procModule);

        // Add a class module.
        VbaModule classModule = new VbaModule
        {
            Name = "MyClass",
            Type = VbaModuleType.ClassModule,
            SourceCode = "Public Sub Greet()\n    MsgBox \"Hello from class module\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(classModule);

        // Add a document module.
        VbaModule docModule = new VbaModule
        {
            Name = "DocumentModule",
            Type = VbaModuleType.DocumentModule,
            SourceCode = "Sub DocumentMacro()\n    MsgBox \"Document module macro\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(docModule);

        // Add a designer module.
        VbaModule designerModule = new VbaModule
        {
            Name = "DesignerModule",
            Type = VbaModuleType.DesignerModule,
            SourceCode = "' Designer module placeholder"
        };
        doc.VbaProject.Modules.Add(designerModule);

        // Remove all modules that are not class modules.
        VbaModuleCollection modules = doc.VbaProject.Modules;
        for (int i = modules.Count - 1; i >= 0; i--)
        {
            VbaModule module = modules[i];
            if (module.Type != VbaModuleType.ClassModule)
                modules.Remove(module);
        }

        // Save the document in a macro-enabled format.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docm");
        doc.Save(outputPath);
    }
}
