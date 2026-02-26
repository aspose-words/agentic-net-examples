using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load an existing document (DOCX or DOCM).
        Document doc = new Document("Input.docx");

        // Verify whether the document contains VBA macros.
        if (doc.HasMacros)
        {
            // Access the VBA project attached to the document.
            VbaProject vbaProject = doc.VbaProject;

            // Output basic information about the VBA project.
            Console.WriteLine($"Project name: {vbaProject.Name}");
            Console.WriteLine($"Code page: {vbaProject.CodePage}");
            Console.WriteLine($"Modules count: {vbaProject.Modules.Count}");

            // ----- Manage VBA project references -----
            // List existing references (if any).
            var references = vbaProject.References;
            Console.WriteLine($"References count: {references.Count}");

            // Example of iterating through references (placeholder for actual reference properties).
            int index = 1;
            foreach (var reference in references)
            {
                // Replace with actual reference details when the reference type is known.
                Console.WriteLine($"Reference #{index++}");
            }

            // Example of adding a new reference.
            // Note: The concrete reference class (e.g., VbaReference) is not shown in the provided API.
            // Uncomment and adjust the following lines when the appropriate class is available.
            // var newReference = new VbaReference();
            // newReference.Name = "MyLibrary";
            // vbaProject.References.Add(newReference);
        }
        else
        {
            Console.WriteLine("The document does not contain any VBA macros.");
        }

        // Save the document, preserving any changes made to the VBA project.
        doc.Save("Output.docx");
    }
}
