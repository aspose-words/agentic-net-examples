using System;
using Aspose.Words;
using Aspose.Words.Vba;

class RemoveVbaReferenceExample
{
    static void Main()
    {
        // Path to the input document that contains a VBA project.
        string inputPath = @"C:\Data\VBA project.docm";

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
        string outputPath = @"C:\Data\VBA project Modified.docm";
        doc.Save(outputPath);
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
