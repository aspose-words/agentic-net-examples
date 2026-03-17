using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the document that contains a VBA project.
        Document doc = new Document("Input.docm");

        // If the document has a VBA project, process its modules.
        if (doc.VbaProject != null)
        {
            VbaModuleCollection modules = doc.VbaProject.Modules;

            // Iterate backwards so that removal does not affect the loop index.
            for (int i = modules.Count - 1; i >= 0; i--)
            {
                VbaModule module = modules[i];

                // Remove the module unless it is a class module.
                if (module.Type != VbaModuleType.ClassModule)
                {
                    modules.Remove(module);
                }
            }
        }

        // Save the document with only class modules remaining.
        doc.Save("Output.docm");
    }
}
