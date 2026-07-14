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

        // Move to the field separator and write a simple EQ argument.
        // This example creates a fraction 1/2 using the "\f" switch.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Ensure the field code is processed before conversion.
        eqField.Update();

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Ensure the conversion succeeded.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
            // Remove the original EQ field from the document.
            eqField.Remove();

            // Set the display type to Inline (default) to keep it inline with the text.
            officeMath.DisplayType = OfficeMathDisplayType.Inline;
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }

        // Save the document to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "InlineEquation.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
