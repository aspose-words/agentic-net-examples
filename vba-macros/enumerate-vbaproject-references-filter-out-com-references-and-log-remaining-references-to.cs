using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaReferenceEnumerator
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a new VBA project and assign it to the document.
            VbaProject project = new VbaProject
            {
                Name = "SampleProject"
            };
            doc.VbaProject = project;

            // Add a simple procedural module so the document is macro-enabled.
            VbaModule module = new VbaModule
            {
                Name = "SampleModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"
            };
            project.Modules.Add(module);

            // Save the document in a macro‑enabled format.
            const string filePath = "SampleDocument.docm";
            doc.Save(filePath);

            // Reload the document (optional, demonstrates loading).
            Document loadedDoc = new Document(filePath);

            // Ensure the document actually contains a VBA project.
            if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
            {
                Console.WriteLine("The document does not contain a VBA project.");
                return;
            }

            // Enumerate the VBA references, filter out COM references, and log the rest.
            VbaReferenceCollection references = loadedDoc.VbaProject.References;

            Console.WriteLine($"Total references: {references.Count}");
            foreach (VbaReference reference in references)
            {
                // COM references are of types Registered, Original, and Control.
                if (reference.Type == VbaReferenceType.Registered ||
                    reference.Type == VbaReferenceType.Original ||
                    reference.Type == VbaReferenceType.Control)
                {
                    // Skip COM references.
                    continue;
                }

                // Remaining references (e.g., Project type) are logged.
                Console.WriteLine($"Reference Type: {reference.Type}, LibId: {reference.LibId}");
            }
        }
    }
}
