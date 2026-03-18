using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source Word document (replace with your actual file path).
        Document doc = new Document("input.docx");

        // Perform a simple find‑and‑replace.
        // The Replace method works on the whole document range and is case‑insensitive by default.
        doc.Range.Replace("_FullName_", "John Doe");

        // Export the updated document to Excel format.
        // SaveFormat.Xlsx specifies the .xlsx workbook output.
        doc.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
