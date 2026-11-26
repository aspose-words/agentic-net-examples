using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaReferenceDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Ensure the document has a VBA project.
            if (doc.VbaProject == null)
            {
                doc.VbaProject = new VbaProject();
            }

            // Record the reference count before removal.
            int countBeforeRemove = doc.VbaProject.References.Count;

            // Remove the first reference if any exist.
            if (countBeforeRemove > 0)
            {
                doc.VbaProject.References.RemoveAt(0);
            }

            // Record the reference count after removal.
            int countAfterRemove = doc.VbaProject.References.Count;

            // Validate that the count decreased by one when a reference was present.
            bool isDecreased = countAfterRemove == countBeforeRemove - 1;

            // Output the validation result.
            Console.WriteLine($"Reference count before remove: {countBeforeRemove}");
            Console.WriteLine($"Reference count after remove: {countAfterRemove}");
            Console.WriteLine($"Reference count decreased: {isDecreased}");

            // Save the document in a macro-enabled format.
            doc.Save("Result.docm", SaveFormat.Docm);
        }
    }
}
