using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaMacroExample
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
                Name = "AsposeVbaProject"
            };
            doc.VbaProject = vbaProject;

            // ---------- Class module ----------
            // Create a class module that contains a public method.
            VbaModule classModule = new VbaModule
            {
                Name = "MyClass",
                Type = VbaModuleType.ClassModule,
                SourceCode = @"Option Explicit

Public Sub SayHello()
    MsgBox ""Hello from MyClass!""
End Sub"
            };

            // Add the class module to the VBA project.
            doc.VbaProject.Modules.Add(classModule);

            // ---------- Procedural module ----------
            // Create a standard module that creates an instance of the class and calls its method.
            VbaModule procModule = new VbaModule
            {
                Name = "MainModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"Option Explicit

Public Sub RunMacro()
    Dim obj As New MyClass
    obj.SayHello
End Sub"
            };

            // Add the procedural module to the VBA project.
            doc.VbaProject.Modules.Add(procModule);

            // Save the document as a macro‑enabled file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VbaProjectClassMacro.docm");
            doc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
