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
            // Create a new blank Word document.
            Document doc = new Document();

            // Create a new VBA project and assign a name.
            VbaProject vbaProject = new VbaProject();
            vbaProject.Name = "MyVbaProject";

            // Attach the VBA project to the document.
            doc.VbaProject = vbaProject;

            // Create a new procedural VBA module.
            VbaModule vbaModule = new VbaModule();
            vbaModule.Name = "MyStandardModule";
            vbaModule.Type = VbaModuleType.ProceduralModule;

            // Assign custom macro code to the module's SourceCode property.
            vbaModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words VBA!""
End Sub
";

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(vbaModule);

            // Define the output file path (macro-enabled .docm format).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VbaProject_CreateMacro.docm");

            // Save the document with macros.
            doc.Save(outputPath, SaveFormat.Docm);

            // Indicate completion.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
