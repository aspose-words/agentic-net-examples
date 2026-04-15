using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will be converted to a real OfficeMath object.
        // Use a simple fraction as the equation content.
        FieldEQ eqField = InsertFieldEQ(builder, @"\f(1,2)");

        // Update the specific field and then all fields in the document.
        eqField.Update();
        doc.UpdateFields();

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Ensure the equation is displayed as a separate paragraph before setting justification.
        officeMath.DisplayType = OfficeMathDisplayType.Display;

        // Set the justification of the equation to center alignment.
        officeMath.Justification = OfficeMathJustification.Center;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Helper method to insert an EQ field with the specified arguments.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the equation arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);
        // Return the cursor to the field start's parent node.
        builder.MoveTo(field.Start.ParentNode);
        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
        return field;
    }
}
