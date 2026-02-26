using Aspose.Words;
using Aspose.Words.Vba;

// Path to the source document that contains VBA macros.
string sourcePath = "Source.docm";

// Path where the cloned document will be saved.
string destinationPath = "Cloned.docm";

// Load the source document.
Document sourceDoc = new Document(sourcePath);

// Create an empty destination document.
Document destinationDoc = new Document();

// Clone the entire VBA project from the source document.
VbaProject clonedProject = sourceDoc.VbaProject.Clone();
destinationDoc.VbaProject = clonedProject;

// OPTIONAL: If you need to replace a specific module with a fresh deep clone,
// remove the automatically cloned module and add a new one.
if (destinationDoc.VbaProject.Modules["Module1"] != null)
{
    VbaModule oldModule = destinationDoc.VbaProject.Modules["Module1"];
    VbaModule newModule = sourceDoc.VbaProject.Modules["Module1"].Clone();
    destinationDoc.VbaProject.Modules.Remove(oldModule);
    destinationDoc.VbaProject.Modules.Add(newModule);
}

// Save the destination document (must be saved as a macro‑enabled format).
destinationDoc.Save(destinationPath, SaveFormat.Docm);
