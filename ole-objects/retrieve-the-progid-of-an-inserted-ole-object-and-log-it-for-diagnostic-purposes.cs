using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

public class Program
{
    public static void Main()
    {
        const string progId = "Shell.Application";

        try
        {
            // Get the COM type from the known ProgID
            Type comType = Type.GetTypeFromProgID(progId);
            if (comType == null)
            {
                Console.WriteLine($"ProgID '{progId}' not found.");
                return;
            }

            // Create an instance of the COM object
            object comObject = Activator.CreateInstance(comType);

            // Retrieve the ProgID by looking it up in the registry using the CLSID (GUID)
            string retrievedProgId = GetProgIdFromGuid(comType.GUID);

            Console.WriteLine($"Inserted OLE object ProgID: {retrievedProgId}");

            // Release the COM object
            Marshal.ReleaseComObject(comObject);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    private static string GetProgIdFromGuid(Guid guid)
    {
        // Registry path: HKEY_CLASSES_ROOT\CLSID\{guid}\ProgID
        string keyPath = $@"CLSID\{{{guid}}}\ProgID";
        using RegistryKey key = Registry.ClassesRoot.OpenSubKey(keyPath);
        if (key != null)
        {
            // The default value of the ProgID subkey holds the ProgID string
            object value = key.GetValue(null);
            if (value is string progId)
            {
                return progId;
            }
        }

        return "Unknown ProgID";
    }
}
