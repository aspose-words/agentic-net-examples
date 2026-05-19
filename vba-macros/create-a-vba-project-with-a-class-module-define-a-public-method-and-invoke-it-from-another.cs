using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaExample
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

            // ----- Class module -----
            // This module defines a class named MyClass with a public method SayHello.
            VbaModule classModule = new VbaModule
            {
                Name = "MyClass",
                Type = VbaModuleType.ClassModule,
                SourceCode = @"
Public Sub SayHello()
    MsgBox ""Hello from MyClass!""
End Sub
"
            };
            doc.VbaProject.Modules.Add(classModule);

            // ----- Procedural module -----
            // This module contains a macro that creates an instance of MyClass and calls SayHello.
            VbaModule proceduralModule = new VbaModule
            {
                Name = "MainModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Public Sub Run()
    Dim obj As New MyClass
    obj.SayHello
End Sub
"
            };
            doc.VbaProject.Modules.Add(proceduralModule);

            // Save the document as a macro‑enabled .docm file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VbaExample.docm");
            doc.Save(outputPath);

            // Load the saved document to verify that the VBA project and modules exist.
            Document loadedDoc = new Document(outputPath);
            Console.WriteLine($"Document has macros: {loadedDoc.HasMacros}");
            Console.WriteLine($"VBA project name: {loadedDoc.VbaProject?.Name}");
            Console.WriteLine($"Number of VBA modules: {loadedDoc.VbaProject?.Modules?.Count}");

            // Output the source code of each module (for demonstration purposes).
            foreach (VbaModule module in loadedDoc.VbaProject.Modules)
            {
                Console.WriteLine($"--- Module: {module.Name} ({module.Type}) ---");
                Console.WriteLine(module.SourceCode?.Trim());
            }
        }
    }
}
