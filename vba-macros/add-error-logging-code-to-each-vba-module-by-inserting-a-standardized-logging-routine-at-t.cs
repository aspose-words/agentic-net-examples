using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaLogging
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Ensure the document has a VBA project.
            VbaProject vbaProject = new VbaProject
            {
                Name = "SampleProject"
            };
            doc.VbaProject = vbaProject;

            // Add a couple of sample VBA modules (procedural modules) with some dummy code.
            VbaModule module1 = new VbaModule
            {
                Name = "ModuleOne",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Sub TestOne()\n    MsgBox \"Hello from ModuleOne\"\nEnd Sub"
            };
            vbaProject.Modules.Add(module1);

            VbaModule module2 = new VbaModule
            {
                Name = "ModuleTwo",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Sub TestTwo()\n    MsgBox \"Hello from ModuleTwo\"\nEnd Sub"
            };
            vbaProject.Modules.Add(module2);

            // Define the standardized logging routine to be inserted at the beginning of each module.
            const string loggingRoutine = 
@"Sub LogError(errMsg As String)
    ' Simple error logging routine
    Debug.Print ""Error: "" & errMsg
End Sub

";

            // Insert the logging routine into every VBA module.
            foreach (VbaModule module in vbaProject.Modules)
            {
                // Guard against null source code.
                string originalSource = module.SourceCode ?? string.Empty;

                // Prepend the logging routine.
                module.SourceCode = loggingRoutine + originalSource;
            }

            // Save the document in a macro-enabled format (.docm).
            const string outputPath = "OutputWithLogging.docm";
            doc.Save(outputPath);
        }
    }
}
