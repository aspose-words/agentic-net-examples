using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToXpsConverter
{
    static void Main()
    {
        // Load the PDF template that contains fields or expressions.
        Document doc = new Document("Template.pdf");

        // Ensure that fields (e.g., MERGEFIELD, FORMFIELD) are evaluated before saving.
        // This is the default behavior (UpdateFields = true), but we set it explicitly for clarity.
        doc.UpdateFields();

        // Create XPS save options. The constructor initializes the object for XPS output.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Optional: customize options, e.g., embed generator name or set color mode.
        // xpsOptions.ExportGeneratorName = true;
        // xpsOptions.ColorMode = ColorMode.Grayscale;

        // Save the document as XPS, which will render the evaluated expressions.
        doc.Save("Result.xps", xpsOptions);
    }
}
