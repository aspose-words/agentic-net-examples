using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class RemoveVbaReferenceExample
{
    static void Main()
    {
        // Path to the input document that contains a VBA project.
        // Use a relative path so the example can run on any machine.
        string inputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VBA project.docm");

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            Console.WriteLine("Place a .docm file named \"VBA project.docm\" in the executable directory and rerun the example.");
            return;
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Access the VBA project.
        VbaProject vbaProject = doc.VbaProject;

        // Get the collection of VBA references.
        VbaReferenceCollection references = vbaProject.References;

        // Store the original count.
        int originalCount = references.Count;
        Console.WriteLine($"Original reference count: {originalCount}");

        // Ensure there is at least one reference to remove.
        if (originalCount == 0)
        {
            Console.WriteLine("No VBA references found to remove.");
            return;
        }

        // Remove the first reference using RemoveAt.
        references.RemoveAt(0);

        // Verify that the count has decreased by one.
        int newCount = references.Count;
        Console.WriteLine($"New reference count after removal: {newCount}");

        if (newCount == originalCount - 1)
            Console.WriteLine("Reference removal confirmed.");
        else
            Console.WriteLine("Reference removal failed.");

        // Save the modified document.
        string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VBA project Modified.docm");
        doc.Save(outputPath);
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
