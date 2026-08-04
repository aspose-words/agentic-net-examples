using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaProjectExample
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
                Name = "ExampleProject"
            };
            doc.VbaProject = vbaProject;

            // ----- Create a class module -----
            VbaModule classModule = new VbaModule
            {
                Name = "MyClass",
                Type = VbaModuleType.ClassModule,
                // Define a public method Hello that shows a message box.
                SourceCode = @"
Public Sub Hello()
    MsgBox ""Hello from MyClass!""
End Sub
"
            };
            // Add the class module to the VBA project.
            doc.VbaProject.Modules.Add(classModule);

            // ----- Create a procedural module that invokes the class method -----
            VbaModule proceduralModule = new VbaModule
            {
                Name = "MainModule",
                Type = VbaModuleType.ProceduralModule,
                // Subroutine that creates an instance of MyClass and calls Hello.
                SourceCode = @"
Sub RunMacro()
    Dim obj As New MyClass
    obj.Hello
End Sub
"
            };
            // Add the procedural module to the VBA project.
            doc.VbaProject.Modules.Add(proceduralModule);

            // Save the document as a macro‑enabled file.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "VbaProjectExample.docm");
            doc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"VBA project created and saved to: {outputPath}");
        }
    }
}
