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
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "AsposeVbaProject";
        doc.VbaProject = vbaProject;

        // ---------- Create a class module ----------
        VbaModule classModule = new VbaModule();
        classModule.Name = "MyClass";
        classModule.Type = VbaModuleType.ClassModule;
        classModule.SourceCode =
@"Option Explicit

Public Sub SayHello()
    MsgBox ""Hello from MyClass!""
End Sub
";
        // Add the class module to the VBA project.
        doc.VbaProject.Modules.Add(classModule);

        // ---------- Create a procedural module that invokes the class method ----------
        VbaModule procModule = new VbaModule();
        procModule.Name = "MainModule";
        procModule.Type = VbaModuleType.ProceduralModule;
        procModule.SourceCode =
@"Option Explicit

Public Sub RunMacro()
    Dim obj As New MyClass
    obj.SayHello
End Sub
";
        // Add the procedural module to the VBA project.
        doc.VbaProject.Modules.Add(procModule);

        // Save the document as a macro‑enabled file.
        const string fileName = "VbaProjectExample.docm";
        doc.Save(fileName);

        // Load the saved document to verify the modules were added.
        Document loadedDoc = new Document(fileName);
        Console.WriteLine("Document has macros: " + loadedDoc.HasMacros);
        Console.WriteLine("VBA Project Name: " + loadedDoc.VbaProject.Name);
        Console.WriteLine("Modules count: " + loadedDoc.VbaProject.Modules.Count);

        // Output the source code of each module.
        foreach (VbaModule module in loadedDoc.VbaProject.Modules)
        {
            Console.WriteLine("--- Module: " + module.Name + " (" + module.Type + ") ---");
            Console.WriteLine(module.SourceCode ?? string.Empty);
        }
    }
}
