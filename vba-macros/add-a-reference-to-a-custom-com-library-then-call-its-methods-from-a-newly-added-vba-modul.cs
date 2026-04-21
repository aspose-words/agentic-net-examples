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

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Set a name for the VBA project.
        doc.VbaProject.Name = "CustomComProject";

        // Create a new procedural VBA module.
        VbaModule vbaModule = new VbaModule
        {
            Name = "CallComModule",
            Type = VbaModuleType.ProceduralModule,
            // VBA code that creates an instance of a custom COM library (late binding) and calls a method.
            SourceCode = @"
Sub CallCustomCom()
    ' Replace ""MyComLib.ProgId"" with the actual ProgID of the COM library.
    Dim comObj As Object
    Set comObj = CreateObject(""MyComLib.ProgId"")
    ' Call a method named ""DoWork"" on the COM object.
    comObj.DoWork
End Sub"
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document in a macro‑enabled format.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomComMacro.docm");
        doc.Save(outputPath, SaveFormat.Docm);

        // Simple verification: output the name of the added module and a snippet of its source.
        Console.WriteLine($"Document saved to: {outputPath}");
        Console.WriteLine($"Added VBA module: {vbaModule.Name}");
        Console.WriteLine("Module source code preview:");
        Console.WriteLine(vbaModule.SourceCode.Trim());
    }
}
