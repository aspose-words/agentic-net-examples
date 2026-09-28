using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field named "AgreementCheckBox",
        // set it to be checked by default, and give it a custom size of 20 points.
        builder.InsertCheckBox("AgreementCheckBox", true, 20);

        // Save the document to the output file.
        doc.Save("CheckboxFormField.docx");
    }
}
