using System;
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

            // Ensure the document has a VBA project; create one if it does not exist.
            VbaProject vbaProject = new VbaProject
            {
                Name = "CustomCOMProject"
            };
            doc.VbaProject = vbaProject;

            // VBA code that creates an instance of a custom COM library (late binding) and calls a method.
            string vbaCode = @"
Sub CallCustomCom()
    Dim comObj As Object
    ' Replace ""MyCustomLib.Class"" with the ProgID of the actual COM library.
    Set comObj = CreateObject(""MyCustomLib.Class"")
    comObj.DoSomething
End Sub
";

            // Create a new procedural VBA module and assign the source code.
            VbaModule vbaModule = new VbaModule
            {
                Name = "CustomComModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = vbaCode
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(vbaModule);

            // Save the document in a macro‑enabled format.
            const string outputPath = "CustomComMacro.docm";
            doc.Save(outputPath);

            // Simple validation output.
            Console.WriteLine($"Document saved to {outputPath}");
            Console.WriteLine($"Has macros: {doc.HasMacros}");
            Console.WriteLine($"Module count: {doc.VbaProject.Modules.Count}");
            Console.WriteLine($"Module \"{vbaModule.Name}\" source code:");
            Console.WriteLine(vbaModule.SourceCode);
        }
    }
}
