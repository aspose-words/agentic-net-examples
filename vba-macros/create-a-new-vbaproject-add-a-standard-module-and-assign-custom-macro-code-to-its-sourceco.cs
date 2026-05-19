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

        // Create a new VBA project and assign a name.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "MyAsposeProject";

        // Attach the VBA project to the document.
        doc.VbaProject = vbaProject;

        // Create a new standard (procedural) VBA module.
        VbaModule vbaModule = new VbaModule();
        vbaModule.Name = "MyStandardModule";
        vbaModule.Type = VbaModuleType.ProceduralModule;

        // Assign custom macro code to the module's SourceCode property.
        vbaModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words VBA!""
End Sub
";

        // Add the module to the VBA project's module collection.
        doc.VbaProject.Modules.Add(vbaModule);

        // Define the output path (current directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VbaProjectExample.docm");

        // Save the document in a macro-enabled format.
        doc.Save(outputPath);
    }
}
