using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaExample
{
    public class Program
    {
        // Standardized logging routine to prepend to each VBA module.
        private const string LoggingRoutine = 
            "'=== Error Logging Routine ===\n" +
            "Sub LogError()\n" +
            "    Debug.Print \"Error at \" & Now\n" +
            "End Sub\n\n";

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

            // Add sample VBA modules.
            AddSampleModule(doc, "Module1", 
                "Sub Test1()\n" +
                "    MsgBox \"Hello from Test1\"\n" +
                "End Sub\n");

            AddSampleModule(doc, "Module2", 
                "Function AddNumbers(a As Integer, b As Integer) As Integer\n" +
                "    AddNumbers = a + b\n" +
                "End Function\n");

            // Insert the logging routine at the beginning of each module.
            foreach (VbaModule module in doc.VbaProject.Modules)
            {
                // Guard against null source code.
                string originalSource = module.SourceCode ?? string.Empty;

                // Prepend the logging routine.
                module.SourceCode = LoggingRoutine + originalSource;
            }

            // Save the document in a macro-enabled format.
            doc.Save("OutputWithLogging.docm");
        }

        // Helper method to create and add a VBA module with given name and source code.
        private static void AddSampleModule(Document doc, string moduleName, string sourceCode)
        {
            VbaModule module = new VbaModule
            {
                Name = moduleName,
                Type = VbaModuleType.ProceduralModule,
                SourceCode = sourceCode
            };
            doc.VbaProject.Modules.Add(module);
        }
    }
}
