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
        vbaProject.Name = "MyCustomProject";
        doc.VbaProject = vbaProject;

        // Create a new procedural VBA module.
        VbaModule vbaModule = new VbaModule();
        vbaModule.Name = "CustomComModule";
        vbaModule.Type = VbaModuleType.ProceduralModule;

        // VBA code that creates an instance of a custom COM library and calls a method.
        // Using a regular string literal with escaped new‑line characters to avoid verbatim‑string issues.
        vbaModule.SourceCode =
            "Sub CallCustomCom()\r\n" +
            "    Dim obj As Object\r\n" +
            "    Set obj = CreateObject(\"MyComLib.MyClass\")\r\n" +
            "    obj.MyMethod\r\n" +
            "End Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document as a macro‑enabled file.
        const string outputPath = "Output.docm";
        doc.Save(outputPath);

        // Simple verification output.
        Console.WriteLine($"Document saved to '{outputPath}'. Module count: {doc.VbaProject.Modules.Count}");
    }
}
