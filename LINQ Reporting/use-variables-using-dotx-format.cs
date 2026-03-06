using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOTX template that contains DOCVARIABLE fields.
        Document doc = new Document("Template.dotx");

        // Set document variables that the DOCVARIABLE fields will display.
        doc.Variables["CompanyName"] = "Acme Corp";
        doc.Variables["Address"] = "123 Main St.";
        doc.Variables["Date"] = DateTime.Today.ToString("d");

        // Refresh all fields so they show the updated variable values.
        doc.UpdateFields();

        // Save the populated document as a regular DOCX file.
        doc.Save("Result.docx");
    }
}
