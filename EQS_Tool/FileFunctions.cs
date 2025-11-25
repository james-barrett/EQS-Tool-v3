using System;
using System.Diagnostics;
using System.IO;

namespace EQS_Tool
{
    internal class FileFunctions
    {
        public static (bool success, string message) Move(string sourcePath, string destinationPath)
        {
            if (File.Exists(destinationPath))
            {
                return (false, $"Destination file '{destinationPath}' already exists.");
            }

            try
            {
                File.Move(sourcePath, destinationPath);
                return (true, $"File moved successfully to '{destinationPath}'.");
            }
            catch (Exception ex)
            {
                return (false, $"Error moving file: {ex.Message}");
            }
        }
    }
}


