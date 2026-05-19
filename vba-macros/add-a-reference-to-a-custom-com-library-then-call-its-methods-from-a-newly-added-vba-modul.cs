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
        VbaProject vbaProject = new VbaProject
        {
            Name = "MyCustomProject"
        };
        doc.VbaProject = vbaProject;

        // Create a new procedural VBA module.
        VbaModule vbaModule = new VbaModule
        {
            Name = "CustomModule",
            Type = VbaModuleType.ProceduralModule,
            // VBA code that calls a method from a custom COM library.
            // The actual COM reference must be added manually in Word; this example shows the intended code.
            SourceCode = @"
Sub CallCustom()
    ' Create an instance of a class from the custom COM library.
    Dim obj As Object
    Set obj = CreateObject(""CustomLib.Class1"")
    obj.DoSomething
End Sub"
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Define the output path for the macro‑enabled document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomMacroDocument.docm");

        // Save the document in a macro‑enabled format.
        doc.Save(outputPath);

        // Load the document again to verify that the module was added.
        Document loadedDoc = new Document(outputPath);
        VbaModule loadedModule = loadedDoc.VbaProject.Modules["CustomModule"];

        // Simple validation: check that the source contains the expected subroutine name.
        bool containsSub = loadedModule?.SourceCode?.Contains("Sub CallCustom()") ?? false;
        Console.WriteLine(containsSub
            ? "VBA module added successfully with expected code."
            : "Failed to add the VBA module or code is missing.");
    }
}
