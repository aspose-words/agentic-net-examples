using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Write("Inline equation: ");

        // Insert an EQ field which will be converted to a real OfficeMath object.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write a simple fraction equation.
        // The EQ switch "\f" creates a fraction; arguments are numerator and denominator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that its internal code is recognized.
        eqField.Update();

        // Return the builder to the paragraph so we can continue adding content after the equation.
        builder.MoveTo(eqField.Start.ParentNode);
        builder.Write(" is displayed inline.");

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Ensure the conversion succeeded.
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the original field and then remove the field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
