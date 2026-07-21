using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "CustomComMacro.docm");

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "CustomComProject"
        };
        doc.VbaProject = vbaProject;

        // Create a new procedural VBA module.
        VbaModule vbaModule = new VbaModule
        {
            Name = "CustomComModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub CallCustomCom()
    Dim obj As Object
    Set obj = CreateObject(""MyCustomLib.MyClass"")
    obj.DoSomething
End Sub"
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document in a macro‑enabled format.
        doc.Save(docPath);

        // Reload the document to verify that the module was added correctly.
        Document loadedDoc = new Document(docPath);
        VbaProject loadedProject = loadedDoc.VbaProject;

        if (loadedProject != null && loadedProject.Modules["CustomComModule"] != null)
        {
            string source = loadedProject.Modules["CustomComModule"].SourceCode ?? string.Empty;
            bool containsCreateObject = source.Contains("CreateObject(\"MyCustomLib.MyClass\")");
            Console.WriteLine(containsCreateObject
                ? "VBA module added with COM reference call."
                : "VBA module added, but COM call not found.");
        }
        else
        {
            Console.WriteLine("Failed to add VBA module.");
        }
    }
}
