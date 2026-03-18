using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load a document that contains content controls (structured document tags).
        Document doc = new Document("InputWithContentControls.docx");

        // Create XpsSaveOptions – this object allows us to customize the XPS export.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Prevent automatic field updates so the original layout of the content controls is kept unchanged.
        xpsOptions.UpdateFields = false;

        // Save the document to XPS format. The content controls are rendered as part of the page,
        // preserving their visual boundaries in the resulting XPS file.
        doc.Save("OutputWithContentControls.xps", xpsOptions);
    }
}
