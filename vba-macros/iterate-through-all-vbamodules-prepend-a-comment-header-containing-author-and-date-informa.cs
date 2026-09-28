using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModuleHeaderAppender
{
    public class Program
    {
        public static void Main()
        {
            // Author information to prepend.
            const string author = "John Doe";

            // Create a new blank document.
            Document doc = new Document();

            // Ensure the document has a VBA project.
            if (doc.VbaProject == null)
            {
                doc.VbaProject = new VbaProject();
            }

            // Access the VBA module collection.
            VbaModuleCollection modules = doc.VbaProject.Modules;

            // If there are no modules, add a sample module for demonstration.
            if (modules.Count == 0)
            {
                string sampleCode =
@"Sub SampleMacro()
    MsgBox ""Hello from VBA!""
End Sub";

                // Create a module using the parameter‑less constructor and set its properties.
                VbaModule sampleModule = new VbaModule();
                sampleModule.Name = "SampleModule";
                sampleModule.SourceCode = sampleCode;

                modules.Add(sampleModule);
            }

            // Iterate through all modules and prepend the comment header.
            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string existingSource = module.SourceCode ?? string.Empty;

                // Build the header comment.
                string header = $"' Author: {author}, Date: {DateTime.Now:yyyy-MM-dd}" + Environment.NewLine;

                // Prepend the header to the existing source.
                module.SourceCode = header + existingSource;
            }

            // Save the document in macro‑enabled format.
            doc.Save("Output.docm");
        }
    }
}
