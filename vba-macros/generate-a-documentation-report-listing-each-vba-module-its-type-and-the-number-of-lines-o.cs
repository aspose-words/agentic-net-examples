using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define paths for the macro-enabled document and the report output.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleMacros.docm");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "VbaModulesReport.txt");

        // -----------------------------------------------------------------
        // Step 1: Create a new blank document.
        Document doc = new Document();

        // Step 2: Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "Aspose.SampleProject";
        doc.VbaProject = vbaProject;

        // Step 3: Add a procedural module.
        VbaModule procModule = new VbaModule();
        procModule.Name = "ProceduralModule";
        procModule.Type = VbaModuleType.ProceduralModule;
        procModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from procedural module!""
End Sub
";
        doc.VbaProject.Modules.Add(procModule);

        // Step 4: Add a class module.
        VbaModule classModule = new VbaModule();
        classModule.Name = "SampleClass";
        classModule.Type = VbaModuleType.ClassModule;
        classModule.SourceCode = @"
Option Explicit

Public Sub Greet()
    MsgBox ""Greetings from class module!""
End Sub
";
        doc.VbaProject.Modules.Add(classModule);

        // Step 5: Add a document module.
        VbaModule docModule = new VbaModule();
        docModule.Name = "DocumentModule";
        docModule.Type = VbaModuleType.DocumentModule;
        docModule.SourceCode = @"
Sub AutoOpen()
    MsgBox ""Document opened!""
End Sub
";
        doc.VbaProject.Modules.Add(docModule);

        // Step 6: Save the document in macro-enabled format.
        doc.Save(docPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // Step 7: Load the saved document (demonstrates loading workflow).
        Document loadedDoc = new Document(docPath);

        // Ensure the document actually contains a VBA project.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Step 8: Prepare the report content.
        using (StreamWriter writer = new StreamWriter(reportPath, false))
        {
            writer.WriteLine("VBA Modules Documentation Report");
            writer.WriteLine($"Generated on: {DateTime.Now}");
            writer.WriteLine(new string('=', 40));
            writer.WriteLine();

            VbaModuleCollection modules = loadedDoc.VbaProject.Modules;
            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Count lines – treat empty source as zero lines.
                int lineCount = string.IsNullOrEmpty(source) ? 0 : source.Split('\n').Length;

                // Write module information.
                writer.WriteLine($"Module Name : {module.Name}");
                writer.WriteLine($"Module Type : {module.Type}");
                writer.WriteLine($"Lines of Code: {lineCount}");
                writer.WriteLine(new string('-', 30));
            }
        }

        // Step 9: Output the report location to the console.
        Console.WriteLine($"VBA modules report generated at: {reportPath}");
    }
}
