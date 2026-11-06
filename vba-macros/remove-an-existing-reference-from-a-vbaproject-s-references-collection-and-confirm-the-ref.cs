using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        doc.VbaProject = new VbaProject();

        // Access the references collection.
        VbaReferenceCollection references = doc.VbaProject.References;

        // Record the initial reference count.
        int initialCount = references.Count;

        // Find the first available reference to remove.
        int indexToRemove = -1;
        for (int i = 0; i < references.Count; i++)
        {
            VbaReference reference = references[i];
            if (reference != null)
            {
                indexToRemove = i;
                break;
            }
        }

        // Remove the reference if one was found.
        if (indexToRemove >= 0)
        {
            references.RemoveAt(indexToRemove);
        }

        // Record the final reference count.
        int finalCount = references.Count;

        // Validate that the count decreased by one.
        bool success = finalCount == initialCount - 1;
        Console.WriteLine(success
            ? "Reference removed successfully. Count decreased."
            : "Reference removal failed.");

        // Save the document in macro-enabled format.
        doc.Save("ReferenceTest.docm", SaveFormat.Docm);
    }
}
