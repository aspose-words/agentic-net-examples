using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "MyVbaProject";
        doc.VbaProject = vbaProject;

        // ---------- Class Module ----------
        // Define a class named MyClass with a public method Greet.
        VbaModule classModule = new VbaModule();
        classModule.Name = "MyClass";
        classModule.Type = VbaModuleType.ClassModule;
        classModule.SourceCode = @"
Public Sub Greet()
    MsgBox ""Hello from MyClass!""
End Sub
";
        // Add the class module to the VBA project.
        doc.VbaProject.Modules.Add(classModule);

        // ---------- Procedural Module ----------
        // Define a standard module that creates an instance of MyClass and calls Greet.
        VbaModule proceduralModule = new VbaModule();
        proceduralModule.Name = "MainModule";
        proceduralModule.Type = VbaModuleType.ProceduralModule;
        proceduralModule.SourceCode = @"
Sub Run()
    Dim obj As New MyClass
    obj.Greet
End Sub
";
        // Add the procedural module to the VBA project.
        doc.VbaProject.Modules.Add(proceduralModule);

        // Save the document as a macro‑enabled file.
        doc.Save("VbaProject_WithClassAndInvocation.docm");
    }
}
