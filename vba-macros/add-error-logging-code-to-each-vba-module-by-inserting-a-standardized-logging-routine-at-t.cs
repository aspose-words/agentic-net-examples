using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a blank document.
            Document doc = new Document();

            // Create a new VBA project and assign it to the document.
            VbaProject vbaProject = new VbaProject();
            vbaProject.Name = "SampleProject";
            doc.VbaProject = vbaProject;

            // Create sample VBA modules with some code.
            VbaModule module1 = new VbaModule();
            module1.Name = "Module1";
            module1.Type = VbaModuleType.ProceduralModule;
            module1.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";

            VbaModule module2 = new VbaModule();
            module2.Name = "Module2";
            module2.Type = VbaModuleType.ProceduralModule;
            module2.SourceCode = "Function Add(a As Integer, b As Integer) As Integer\n    Add = a + b\nEnd Function";

            // Add the modules to the VBA project.
            doc.VbaProject.Modules.Add(module1);
            doc.VbaProject.Modules.Add(module2);

            // Standardized logging routine to prepend to each module.
            const string loggingRoutine = "' Standard error logging routine\nSub LogError(errMsg As String)\n    Debug.Print \"Error: \" & errMsg\nEnd Sub\n\n";

            // Insert the logging routine at the beginning of each module's source code.
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                string originalSource = module.SourceCode ?? string.Empty;
                if (!originalSource.StartsWith("' Standard error logging routine"))
                {
                    module.SourceCode = loggingRoutine + originalSource;
                }
            }

            // Save the document as a macro-enabled .docm file.
            string outputPath = "OutputWithLogging.docm";
            doc.Save(outputPath);
        }
    }
}
