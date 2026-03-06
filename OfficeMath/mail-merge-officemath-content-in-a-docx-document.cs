using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD where the OfficeMath equation will be placed.
        builder.InsertField("MERGEFIELD Equation \\* MERGEFORMAT");
        builder.Writeln();

        // Attach a custom callback that will replace the merge field with an OfficeMath object.
        doc.MailMerge.FieldMergingCallback = new OfficeMathMergingCallback();

        // Execute the mail merge. The actual value is not used because the callback inserts the equation.
        doc.MailMerge.Execute(new[] { "Equation" }, new object[] { "" });

        // Save the resulting document.
        doc.Save("OfficeMathMailMerge.docx");
    }

    // Implements IFieldMergingCallback to insert OfficeMath during mail merge.
    private class OfficeMathMergingCallback : IFieldMergingCallback
    {
        // Called for each MERGEFIELD encountered.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Process only the specific field we are interested in.
            if (args.DocumentFieldName.Equals("Equation", StringComparison.OrdinalIgnoreCase))
            {
                // Position the builder at the merge field.
                DocumentBuilder cb = new DocumentBuilder(args.Document);
                cb.MoveToMergeField(args.DocumentFieldName);

                // Insert a simple linear equation (a + b = c) using the EQ field syntax.
                // Aspose.Words renders this as an OfficeMath (OMML) object.
                cb.InsertField(@"EQ \o(\a(a),\a(b)) = \a(c)");

                // Prevent the default text insertion for this field.
                args.Text = string.Empty;
            }
        }

        // No image handling required for this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
