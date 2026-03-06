// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Data;
using System.Data.SqlClient;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        string inputPath = "input.docx";
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Retrieve the image bytes from the database (example record id = 123).
        byte[] imageBytes = GetImageFromDatabase(123);

        // Insert the image at the current cursor position.
        builder.InsertImage(imageBytes);

        // Configure PNG save options (optional: set resolution, etc.).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        pngOptions.Resolution = 300; // 300 DPI for higher quality.

        // Save the document as a PNG image (renders the first page).
        string outputPath = "output.png";
        doc.Save(outputPath, pngOptions);
    }

    // Helper method to fetch an image stored as a varbinary column in the database.
    static byte[] GetImageFromDatabase(int imageId)
    {
        // TODO: Replace with your actual connection string.
        string connectionString = "Data Source=SERVER;Initial Catalog=DB;Integrated Security=True;";

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            using (SqlCommand command = new SqlCommand(
                "SELECT ImageData FROM Images WHERE Id = @Id", connection))
            {
                command.Parameters.Add("@Id", SqlDbType.Int).Value = imageId;

                object result = command.ExecuteScalar();

                if (result != null && result != DBNull.Value)
                    return (byte[])result;

                throw new InvalidOperationException("Image not found in the database.");
            }
        }
    }
}
