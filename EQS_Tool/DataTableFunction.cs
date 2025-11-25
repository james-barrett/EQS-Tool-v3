using Sylvan.Data.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using OpenXMLSheet = DocumentFormat.OpenXml.Spreadsheet;


namespace EQS_Tool
{
    public class DataTableFunctions
    {
        public static DataTable LoadDataTable(string excelFilePath, string logFilePath, string timestamp)
        {
            if (!File.Exists(excelFilePath))
            {
                string message = "Missing dependency address file DB, exiting.....";
                Logger.Log(message, logFilePath, LogLevel.Info);
                Helpers.CleanExit();
            }

            try
            {
                DataTable addressTable = new DataTable();
                using (var excelReader = ExcelDataReader.Create(excelFilePath))
                {
                    addressTable.Load(excelReader);
                }
                string message = "Address data loaded successfully.";
                Logger.Log(message, logFilePath, LogLevel.Info);
                return addressTable;
            }
            catch (Exception ex)
            {
                string message = $"Error loading address data, quitting. Exception: {ex.Message}.";
                Logger.Log(message, logFilePath, LogLevel.Info);
                Helpers.CleanExit();
                return null;
            }
        }
        public static string GetAddressFromDataTable(DataTable addressTable, string uprn, string logFilePath, string timestamp)
        {
            const string columnName = "[Property Reference]";
            string address = "UPRN_ERROR";

            try
            {
                DataRow[] matchingRows = addressTable.Select($"{columnName} = '{uprn}'");

                if (matchingRows.Length > 0)
                {
                    address = matchingRows[0]["Property Address"].ToString();
                }
                else
                {
                    string message = $"No matching address found for UPRN: {uprn}";
                    Logger.Log(message, logFilePath, LogLevel.Info);
                }
            }
            catch (Exception ex)
            {
                string message = $"Error retrieving address for UPRN: {uprn}. Exception: {ex.Message}";
                Logger.Log(message, logFilePath, LogLevel.Info);
            }

            return address;
        }
        public static string GetUPRNFromDataTable(DataTable addressTable, string address, string logFilePath, string timestamp)
        {
            const string columnName = "Property Address";
            string uprn = "UPRN_ERROR";

            Console.WriteLine(address);

            try
            {
                //DataRow[] matchingRows = addressTable.Select($"{columnName} = '{address}'");
                string normalizedAddress = NormalizeString(address);

                Console.WriteLine(normalizedAddress);

                //DataRow[] matchingRows = addressTable.Select()
                //    .Where(row => NormalizeString(row[columnName].ToString())
                //   .Equals(normalizedAddress, StringComparison.OrdinalIgnoreCase))
                //   .ToArray();

                // Normalize and partially match addresses from DataTable
                DataRow[] matchingRows = addressTable.Select()
                    .Where(row => NormalizeString(row[columnName].ToString())
                    .Contains(normalizedAddress, StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                if (matchingRows.Length > 0)
                {
                    uprn = matchingRows[0]["Property Reference"].ToString();
                }
                else
                {
                    string message = $"No matching address found for UPRN: {uprn}";
                    Logger.Log(message, logFilePath, LogLevel.Info);
                }
            }
            catch (Exception ex)
            {
                string message = $"Error retrieving address for UPRN: {uprn}. Exception: {ex.Message}";
                Logger.Log(message, logFilePath, LogLevel.Info);
            }

            Console.WriteLine(uprn);

            return uprn;
        }
        public static string NormalizeString(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            return Regex.Replace(input.Replace(",", "").Trim(), @"\s+", " ");
        }
    }
}
