using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

using OpenXMLSheet = DocumentFormat.OpenXml.Spreadsheet;

namespace EQS_Tool
{
    public class GenerateAuditData
    {
        public static int HeaderIndex(string[] headers, string searchText)
        {
            int index = Array.FindIndex(headers, h => h.Trim().Contains(searchText, StringComparison.OrdinalIgnoreCase));
            return index;
        }

        public static void PrecheckCsv(string path, string line)
        {
            string[] headerFields = File.ReadLines(path).First().Split(',');
            int expectedColumnCount = headerFields.Length;

            string[] fields = line.Split(',');

            if (fields.Length < expectedColumnCount)
            {
                Console.WriteLine($"Line has {fields.Length} fields, expected {expectedColumnCount}:\n{line}\n");
            }
        }

        public static string[] ParseCsvLine(string line)
        {
            var fields = new List<string>();
            bool inQuotes = false;
            var currentField = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    // Check for escaped quotes ("")
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++; // Skip the next quote
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }

            // Add the last field
            fields.Add(currentField.ToString());
            return fields.ToArray();
        }

        public static void ProcessAuditData(string path, string[] inputAuditDate, string filePath, List<string> eqsTeam, string month, string year)
        {
            int count = 0;

            using var reader = new StreamReader(path);
            string? line;

            var audits = new List<string>();
            var post = new List<string>();
            var wig = new List<string>();


            var employee = new List<string>();
            var subcon = new List<string>();

            reader.ReadLine();

            string headerLine = File.ReadLines(path).First();
            string[] headers = ParseCsvLine(headerLine);


            //post columns
            int EQSIndex = GenerateAuditData.HeaderIndex(headers, "Please choose your name: Answer");
            int DeptIndex = GenerateAuditData.HeaderIndex(headers, "Which area of the business are they working for?: Answer");
            int GainedAccessIndex = GenerateAuditData.HeaderIndex(headers, "Did you gain access to the property today?: Answer");
            int ResultIndex = GenerateAuditData.HeaderIndex(headers, "Overall Inspection Result: Answer");
            int SummaryIndex = GenerateAuditData.HeaderIndex(headers, "Summary of issues to resolve: Answer");
            int EmpSubIndex = GenerateAuditData.HeaderIndex(headers, "Is this an inspection of a GP Electrician or Contractor?: Answer");

            //both
            int auditorIndex = GenerateAuditData.HeaderIndex(headers, "Auditor's Name");
            int auditDateIndex = GenerateAuditData.HeaderIndex(headers, "Audit Date");
            int auditTypeIndex = GenerateAuditData.HeaderIndex(headers, "Question Set Name");
            int AddressIndex = GenerateAuditData.HeaderIndex(headers, "Description / Site / Building Address / Resident Address");
            int auditeeIndex = GenerateAuditData.HeaderIndex(headers, "Auditee Name");
            int RefIndex = GenerateAuditData.HeaderIndex(headers, "Reference");


            while ((line = reader.ReadLine()) != null)
            {
                var fields = ParseCsvLine(line);

                string auditType;
                string auditDateStr;

                // With this improved version:
                if (fields.Length == 0) continue;

                auditDateStr = fields[auditDateIndex].Trim();

                if (!DateTime.TryParse(auditDateStr, out DateTime auditDate))
                {
                    continue;
                }

                string inputMonth = inputAuditDate[1];
                string inputYear = inputAuditDate[0];

                auditType = fields[auditTypeIndex].Trim();

                if (auditDate.Month.ToString() == inputMonth && auditDate.Year.ToString() == inputYear)
                {
                    if (auditType.Contains("Electrical", StringComparison.OrdinalIgnoreCase))
                    {
                        if (auditType.Contains("Post", StringComparison.OrdinalIgnoreCase))
                        {
                            string access = string.Equals(fields[GainedAccessIndex], "no", StringComparison.OrdinalIgnoreCase)
                                       ? "No"
                                       : "Yes";

                            string auditData = $"{auditDateStr}:" +
                                    $"{fields[auditorIndex]}:" +
                                    $"{fields[auditeeIndex]}:" +
                                    $"{fields[RefIndex]}:" +
                                    $"{fields[EmpSubIndex]}:" +
                                    $"{fields[DeptIndex]}:" +
                                    $"{access}:" +
                                    $"{fields[AddressIndex]}:" +
                                    $"{fields[SummaryIndex]}:" +
                                    $"{fields[ResultIndex]}:" +
                                    $"{(line.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase) ? "unsatisfactory" : "satisfactory")}";


                            if (fields[EmpSubIndex].Contains("Subcontractor", StringComparison.OrdinalIgnoreCase))
                            {
                                subcon.Add(auditData);
                            }
                            else
                            {
                                employee.Add(auditData);
                            }



                            if (fields[DeptIndex].Contains("Fixed Wire Test", StringComparison.OrdinalIgnoreCase))
                            {
                                wig.Add(auditData);
                            }
                            else
                            {
                                post.Add(auditData);

                            }
                        }
                    }
                    else if (auditType.Contains("Electrician", StringComparison.OrdinalIgnoreCase))
                    {
                        audits.Add($"{auditDateStr}:{fields[auditorIndex]}:{fields[auditeeIndex]}:{fields[RefIndex]}");
                    }
                    else
                    {
                        continue;
                    }
                }
            }

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new OpenXMLSheet.Workbook();

                var styles = AddStyles(workbookPart);

                // Create the main worksheet
                WorksheetPart worksheetpart = workbookPart.AddNewPart<WorksheetPart>();

                var columns = new OpenXMLSheet.Columns(
                    new OpenXMLSheet.Column { Min = 1, Max = 1, Width = 30, CustomWidth = true },
                    new OpenXMLSheet.Column { Min = 2, Max = 2, Width = 20, CustomWidth = true },
                    new OpenXMLSheet.Column { Min = 3, Max = 3, Width = 30, CustomWidth = true },
                    new OpenXMLSheet.Column { Min = 4, Max = 4, Width = 30, CustomWidth = true }
                );

                // Create SheetData once
                var sheetData = new OpenXMLSheet.SheetData();
                worksheetpart.Worksheet = new OpenXMLSheet.Worksheet(columns, sheetData);

                // Create Sheets collection
                OpenXMLSheet.Sheets sheets = workbookPart.Workbook.AppendChild(new OpenXMLSheet.Sheets());

                // Add main sheet
                OpenXMLSheet.Sheet sheet = new OpenXMLSheet.Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetpart),
                    SheetId = 1,
                    Name = $"EQS Audit Results {month} {year}"
                };
                sheets.Append(sheet);

                // Create header row with proper RowIndex
                OpenXMLSheet.Row headerRow = new OpenXMLSheet.Row() { RowIndex = 1 };
                headerRow.Append(
                    CreateTextCell("A1", "Name", 2),
                    CreateTextCell("B1", "H&S Audits", 2),
                    CreateTextCell("C1", "Employee Post Inspections", 2),
                    //CreateTextCell("D1", "Wiggets Inspections", 2)
                    CreateTextCell("D1", "Sub-con Post Inspections", 2)
                );
                sheetData.Append(headerRow);

                // Add data rows
                uint rowIndex = 2;
                foreach (string eqs in eqsTeam)
                {


                    int auditCount = audits.Count(x => x.ToLower().Contains(eqs.Split(" ").Last().ToLower()));
                    //int w = wig.Count(x => x.ToLower().Contains(eqs.Split(" ").Last().ToLower()));
                    int subconCount = subcon.Count(x => x.ToLower().Contains(eqs.Split(" ").Last().ToLower()));
                    //int p = post.Count(x => x.ToLower().Contains(eqs.Split(" ").Last().ToLower()));
                    int employeeCount = employee.Count(x => x.ToLower().Contains(eqs.Split(" ").Last().ToLower()));

                    // Create row with proper RowIndex
                    var dataRow = new OpenXMLSheet.Row() { RowIndex = rowIndex };
                    dataRow.Append(
                        CreateTextCell($"A{rowIndex}", eqs.ToUpper(), 1),
                        CreateIntegerCell($"B{rowIndex}", auditCount, 1),
                        CreateIntegerCell($"C{rowIndex}", employeeCount, 1),
                        CreateIntegerCell($"D{rowIndex}", subconCount, 1)
                    );
                    sheetData.Append(dataRow);
                    rowIndex++;
                }

                // Save after main sheet creation
                workbookPart.Workbook.Save();

                // Create Wiggets Audits sheet at index 2
                var wiggetsSheetPart = workbookPart.AddNewPart<WorksheetPart>();

                var wiggetsColumns = new OpenXMLSheet.Columns(
                    new OpenXMLSheet.Column { Min = 1, Max = 1, Width = 20, CustomWidth = true }, // Status & Date
                    new OpenXMLSheet.Column { Min = 2, Max = 2, Width = 15, CustomWidth = true }, // Count & REF
                    new OpenXMLSheet.Column { Min = 3, Max = 3, Width = 20, CustomWidth = true }, // Percentage & EQS
                    new OpenXMLSheet.Column { Min = 4, Max = 4, Width = 20, CustomWidth = true }, // OP
                    new OpenXMLSheet.Column { Min = 5, Max = 5, Width = 20, CustomWidth = true },  // Status
                    new OpenXMLSheet.Column { Min = 6, Max = 6, Width = 15, CustomWidth = true },  // Access
                    new OpenXMLSheet.Column { Min = 7, Max = 7, Width = 40, CustomWidth = true }, // Address
                    new OpenXMLSheet.Column { Min = 8, Max = 8, Width = 40, CustomWidth = true }  // Summary
                );

                var wiggetsSheetData = new OpenXMLSheet.SheetData();
                wiggetsSheetPart.Worksheet = new OpenXMLSheet.Worksheet(wiggetsColumns, wiggetsSheetData);

                // Add Wiggets Audits sheet to sheets collection
                sheets.Append(new OpenXMLSheet.Sheet
                {
                    Id = workbookPart.GetIdOfPart(wiggetsSheetPart),
                    SheetId = 2,
                    Name = "Wiggets Audits"
                });

                int unsatisfactory = wig.Count(x =>
                    x.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase));

                // Total wigget audits (same calculation as int w in the main loop)
                int totalWiggets = wig.Count();

                // Calculate satisfactory and unsatisfactory counts
                int satisfactory = totalWiggets - unsatisfactory;

                // Create header row for Wiggets sheet
                var wiggetsHeaderRow = new OpenXMLSheet.Row() { RowIndex = 1 };
                wiggetsHeaderRow.Append(
                    CreateTextCell("A1", "Status", 2),
                    CreateTextCell("B1", "Count", 2),
                    CreateTextCell("C1", "Percentage", 2)
                );
                wiggetsSheetData.Append(wiggetsHeaderRow);

                // Calculate percentages
                double satPercentage = totalWiggets == 0 ? 0 : (satisfactory / (double)totalWiggets) * 100;
                double unsatPercentage = totalWiggets == 0 ? 0 : (unsatisfactory / (double)totalWiggets) * 100;

                // Add satisfactory row
                var satisfactoryRow = new OpenXMLSheet.Row() { RowIndex = 2 };
                satisfactoryRow.Append(
                    CreateTextCell("A2", "Satisfactory", 1),
                    CreateIntegerCell("B2", satisfactory, 1),
                    CreateTextCell("C2", satPercentage.ToString("F1") + "%", 1)
                );
                wiggetsSheetData.Append(satisfactoryRow);

                // Add unsatisfactory row
                var unsatisfactoryRow = new OpenXMLSheet.Row() { RowIndex = 3 };
                unsatisfactoryRow.Append(
                    CreateTextCell("A3", "Unsatisfactory", 1),
                    CreateIntegerCell("B3", unsatisfactory, 1),
                    CreateTextCell("C3", unsatPercentage.ToString("F1") + "%", 1)
                );
                wiggetsSheetData.Append(unsatisfactoryRow);

                // Add total row
                var totalRow = new OpenXMLSheet.Row() { RowIndex = 4 };
                totalRow.Append(
                    CreateTextCell("A4", "Total", 2),
                    CreateIntegerCell("B4", totalWiggets, 2),
                    CreateTextCell("C4", "", 2)
                );
                wiggetsSheetData.Append(totalRow);

                // Add empty row for spacing
                var emptyRow = new OpenXMLSheet.Row() { RowIndex = 5 };
                wiggetsSheetData.Append(emptyRow);

                // Add detailed entries header
                var detailHeaderRow = new OpenXMLSheet.Row() { RowIndex = 6 };
                detailHeaderRow.Append(
                    CreateTextCell("A6", "Date", 2),
                    CreateTextCell("B6", "REF", 2),
                    CreateTextCell("C6", "EQS", 2),
                    CreateTextCell("D6", "OP", 2),
                    CreateTextCell("E6", "Status", 2),
                    CreateTextCell("F6", "Access", 2),
                    CreateTextCell("G6", "Address", 2),
                    CreateTextCell("H6", "Summary", 2)
                );
                wiggetsSheetData.Append(detailHeaderRow);

                // Add all Fixed Wire Testing entries
                uint detailRowIndex = 7;

                foreach (var entry in wig)
                {
                    var parts = entry.Split(':', StringSplitOptions.TrimEntries);

                    if (parts.Length >= 4)
                    {
                        // Determine status from the entry content
                        string status = "Unknown";
                        if (entry.Contains("satisfactory", StringComparison.OrdinalIgnoreCase))
                        {
                            if (entry.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase))
                                status = "Unsatisfactory"; // "unsatisfactory" takes precedence if both exist
                            else
                                status = "Satisfactory";
                        }
                        else if (entry.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase))
                        {
                            status = "Unsatisfactory";
                        }

                        var detailRow = new OpenXMLSheet.Row() { RowIndex = detailRowIndex };
                        detailRow.Append(
                            CreateTextCell($"A{detailRowIndex}", parts[0], 1), // Date
                            CreateTextCell($"B{detailRowIndex}", parts[3], 1), // REF
                            CreateTextCell($"C{detailRowIndex}", parts[1], 1), // EQS
                            CreateTextCell($"D{detailRowIndex}", parts[2], 1), // OP                            
                            CreateTextCell($"E{detailRowIndex}", status, 1),     // Status
                            CreateTextCell($"F{detailRowIndex}", parts[6], 1),  // Access
                            CreateTextCell($"G{detailRowIndex}", parts[7], 1),  // Add
                            CreateTextCell($"H{detailRowIndex}", parts[8], 1)  // Summary
                        );
                        wiggetsSheetData.Append(detailRow);
                        detailRowIndex++;
                    }
                }

                // Create employee post sheet at index 3
                var employeeSheetPart = workbookPart.AddNewPart<WorksheetPart>();

                var employeeColumns = new OpenXMLSheet.Columns(
                    new OpenXMLSheet.Column { Min = 1, Max = 1, Width = 20, CustomWidth = true }, // Status & Date
                    new OpenXMLSheet.Column { Min = 2, Max = 2, Width = 15, CustomWidth = true }, // Count & REF
                    new OpenXMLSheet.Column { Min = 3, Max = 3, Width = 20, CustomWidth = true }, // Percentage & EQS
                    new OpenXMLSheet.Column { Min = 4, Max = 4, Width = 20, CustomWidth = true }, // OP
                    new OpenXMLSheet.Column { Min = 5, Max = 5, Width = 20, CustomWidth = true },  // Status
                    new OpenXMLSheet.Column { Min = 6, Max = 6, Width = 15, CustomWidth = true },  // Access
                    new OpenXMLSheet.Column { Min = 7, Max = 7, Width = 40, CustomWidth = true }, // Address
                    new OpenXMLSheet.Column { Min = 8, Max = 8, Width = 40, CustomWidth = true }  // Summary
                );

                var employeeSheetData = new OpenXMLSheet.SheetData();
                employeeSheetPart.Worksheet = new OpenXMLSheet.Worksheet(employeeColumns, employeeSheetData);

                sheets.Append(new OpenXMLSheet.Sheet
                {
                    Id = workbookPart.GetIdOfPart(employeeSheetPart),
                    SheetId = 2,
                    Name = "Employee Post Inspections"
                });

                int employeeUnsatisfactory = employee.Count(x =>
                    x.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase));

                // Total wigget audits (same calculation as int w in the main loop)
                int totalEmployee = employee.Count();

                // Calculate satisfactory and unsatisfactory counts
                int employeeSatisfactory = totalEmployee - employeeUnsatisfactory;

                // Create header row for Wiggets sheet
                var employeeHeaderRow = new OpenXMLSheet.Row() { RowIndex = 1 };
                employeeHeaderRow.Append(
                    CreateTextCell("A1", "Status", 2),
                    CreateTextCell("B1", "Count", 2),
                    CreateTextCell("C1", "Percentage", 2)
                );
                employeeSheetData.Append(employeeHeaderRow);

                // Calculate percentages
                double employeeSatPercentage = totalEmployee == 0 ? 0 : (employeeSatisfactory / (double)totalEmployee) * 100;
                double employeeUnsatPercentage = totalEmployee == 0 ? 0 : (employeeUnsatisfactory / (double)totalEmployee) * 100;

                // Add satisfactory row
                var employeeSatisfactoryRow = new OpenXMLSheet.Row() { RowIndex = 2 };
                employeeSatisfactoryRow.Append(
                    CreateTextCell("A2", "Satisfactory", 1),
                    CreateIntegerCell("B2", employeeSatisfactory, 1),
                    CreateTextCell("C2", employeeSatPercentage.ToString("F1") + "%", 1)
                );
                employeeSheetData.Append(employeeSatisfactoryRow);

                // Add unsatisfactory row
                var employeeUnsatisfactoryRow = new OpenXMLSheet.Row() { RowIndex = 3 };
                employeeUnsatisfactoryRow.Append(
                    CreateTextCell("A3", "Unsatisfactory", 1),
                    CreateIntegerCell("B3", employeeUnsatisfactory, 1),
                    CreateTextCell("C3", employeeUnsatPercentage.ToString("F1") + "%", 1)
                );
                employeeSheetData.Append(employeeUnsatisfactoryRow);

                // Add total row
                var employeeTotalRow = new OpenXMLSheet.Row() { RowIndex = 4 };
                employeeTotalRow.Append(
                    CreateTextCell("A4", "Total", 2),
                    CreateIntegerCell("B4", totalEmployee, 2),
                    CreateTextCell("C4", "", 2)
                );
                employeeSheetData.Append(employeeTotalRow);

                // Add empty row for spacing
                var employeeEmptyRow = new OpenXMLSheet.Row() { RowIndex = 5 };
                employeeSheetData.Append(employeeEmptyRow);

                // Add detailed entries header
                var employeeDetailHeaderRow = new OpenXMLSheet.Row() { RowIndex = 6 };
                employeeDetailHeaderRow.Append(
                    CreateTextCell("A6", "Date", 2),
                    CreateTextCell("B6", "REF", 2),
                    CreateTextCell("C6", "EQS", 2),
                    CreateTextCell("D6", "OP", 2),
                    CreateTextCell("E6", "Status", 2),
                    CreateTextCell("F6", "Access", 2),
                    CreateTextCell("G6", "Address", 2),
                    CreateTextCell("H6", "Summary", 2)
                );
                employeeSheetData.Append(employeeDetailHeaderRow);

                // Add all Fixed Wire Testing entries
                uint employeeDetailRowIndex = 7;

                foreach (var entry in employee)
                {
                    var parts = entry.Split(':', StringSplitOptions.TrimEntries);

                    if (parts.Length >= 4)
                    {
                        // Determine status from the entry content
                        string status = "Unknown";
                        if (entry.Contains("satisfactory", StringComparison.OrdinalIgnoreCase))
                        {
                            if (entry.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase))
                                status = "Unsatisfactory"; // "unsatisfactory" takes precedence if both exist
                            else
                                status = "Satisfactory";
                        }
                        else if (entry.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase))
                        {
                            status = "Unsatisfactory";
                        }

                        var employeeDetailRow = new OpenXMLSheet.Row() { RowIndex = employeeDetailRowIndex };
                        employeeDetailRow.Append(
                            CreateTextCell($"A{employeeDetailRowIndex}", parts[0], 1), // Date
                            CreateTextCell($"B{employeeDetailRowIndex}", parts[3], 1), // REF
                            CreateTextCell($"C{employeeDetailRowIndex}", parts[1], 1), // EQS
                            CreateTextCell($"D{employeeDetailRowIndex}", parts[2], 1), // OP                            
                            CreateTextCell($"E{employeeDetailRowIndex}", status, 1),     // Status
                            CreateTextCell($"F{employeeDetailRowIndex}", parts[6], 1),  // Access
                            CreateTextCell($"G{employeeDetailRowIndex}", parts[7], 1),  // Add
                            CreateTextCell($"H{employeeDetailRowIndex}", parts[8], 1)  // Summary
                        );
                        employeeSheetData.Append(employeeDetailRow);
                        employeeDetailRowIndex++;
                    }
                }







                // Create sub con post sheet at index 4
                var subconSheetPart = workbookPart.AddNewPart<WorksheetPart>();

                var subconColumns = new OpenXMLSheet.Columns(
                    new OpenXMLSheet.Column { Min = 1, Max = 1, Width = 20, CustomWidth = true }, // Status & Date
                    new OpenXMLSheet.Column { Min = 2, Max = 2, Width = 15, CustomWidth = true }, // Count & REF
                    new OpenXMLSheet.Column { Min = 3, Max = 3, Width = 20, CustomWidth = true }, // Percentage & EQS
                    new OpenXMLSheet.Column { Min = 4, Max = 4, Width = 20, CustomWidth = true }, // OP
                    new OpenXMLSheet.Column { Min = 5, Max = 5, Width = 20, CustomWidth = true },  // Status
                    new OpenXMLSheet.Column { Min = 6, Max = 6, Width = 15, CustomWidth = true },  // Access
                    new OpenXMLSheet.Column { Min = 7, Max = 7, Width = 40, CustomWidth = true }, // Address
                    new OpenXMLSheet.Column { Min = 8, Max = 8, Width = 40, CustomWidth = true }  // Summary
                );

                var subconSheetData = new OpenXMLSheet.SheetData();
                subconSheetPart.Worksheet = new OpenXMLSheet.Worksheet(subconColumns, subconSheetData);

                sheets.Append(new OpenXMLSheet.Sheet
                {
                    Id = workbookPart.GetIdOfPart(subconSheetPart),
                    SheetId = 2,
                    Name = "Sub-Con Post Inspections"
                });

                int subconUnsatisfactory = subcon.Count(x =>
                    x.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase));

                // Total wigget audits (same calculation as int w in the main loop)
                int totalSubcon = subcon.Count();

                // Calculate satisfactory and unsatisfactory counts
                int subconSatisfactory = totalSubcon - subconUnsatisfactory;

                // Create header row for Wiggets sheet
                var subconHeaderRow = new OpenXMLSheet.Row() { RowIndex = 1 };
                subconHeaderRow.Append(
                    CreateTextCell("A1", "Status", 2),
                    CreateTextCell("B1", "Count", 2),
                    CreateTextCell("C1", "Percentage", 2)
                );
                subconSheetData.Append(subconHeaderRow);

                // Calculate percentages
                double subconSatPercentage = totalSubcon == 0 ? 0 : (subconSatisfactory / (double)totalSubcon) * 100;
                double subconUnsatPercentage = totalSubcon == 0 ? 0 : (subconUnsatisfactory / (double)totalSubcon) * 100;

                // Add satisfactory row
                var subconSatisfactoryRow = new OpenXMLSheet.Row() { RowIndex = 2 };
                subconSatisfactoryRow.Append(
                    CreateTextCell("A2", "Satisfactory", 1),
                    CreateIntegerCell("B2", subconSatisfactory, 1),
                    CreateTextCell("C2", subconSatPercentage.ToString("F1") + "%", 1)
                );
                subconSheetData.Append(subconSatisfactoryRow);

                // Add unsatisfactory row
                var subconUnsatisfactoryRow = new OpenXMLSheet.Row() { RowIndex = 3 };
                subconUnsatisfactoryRow.Append(
                    CreateTextCell("A3", "Unsatisfactory", 1),
                    CreateIntegerCell("B3", subconUnsatisfactory, 1),
                    CreateTextCell("C3", subconUnsatPercentage.ToString("F1") + "%", 1)
                );
                subconSheetData.Append(subconUnsatisfactoryRow);

                // Add total row
                var subconTotalRow = new OpenXMLSheet.Row() { RowIndex = 4 };
                subconTotalRow.Append(
                    CreateTextCell("A4", "Total", 2),
                    CreateIntegerCell("B4", totalSubcon, 2),
                    CreateTextCell("C4", "", 2)
                );
                subconSheetData.Append(subconTotalRow);

                // Add empty row for spacing
                var subconEmptyRow = new OpenXMLSheet.Row() { RowIndex = 5 };
                subconSheetData.Append(subconEmptyRow);

                // Add detailed entries header
                var subconDetailHeaderRow = new OpenXMLSheet.Row() { RowIndex = 6 };
                detailHeaderRow.Append(
                    CreateTextCell("A6", "Date", 2),
                    CreateTextCell("B6", "REF", 2),
                    CreateTextCell("C6", "EQS", 2),
                    CreateTextCell("D6", "OP", 2),
                    CreateTextCell("E6", "Status", 2),
                    CreateTextCell("F6", "Access", 2),
                    CreateTextCell("G6", "Address", 2),
                    CreateTextCell("H6", "Summary", 2)
                );
                subconSheetData.Append(subconDetailHeaderRow);

                // Add all Fixed Wire Testing entries
                uint subconDetailRowIndex = 7;

                foreach (var entry in subcon)
                {
                    var parts = entry.Split(':', StringSplitOptions.TrimEntries);

                    if (parts.Length >= 4)
                    {
                        // Determine status from the entry content
                        string status = "Unknown";
                        if (entry.Contains("satisfactory", StringComparison.OrdinalIgnoreCase))
                        {
                            if (entry.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase))
                                status = "Unsatisfactory"; // "unsatisfactory" takes precedence if both exist
                            else
                                status = "Satisfactory";
                        }
                        else if (entry.Contains("unsatisfactory", StringComparison.OrdinalIgnoreCase))
                        {
                            status = "Unsatisfactory";
                        }

                        var subconDetailRow = new OpenXMLSheet.Row() { RowIndex = subconDetailRowIndex };
                        subconDetailRow.Append(
                            CreateTextCell($"A{subconDetailRowIndex}", parts[0], 1), // Date
                            CreateTextCell($"B{subconDetailRowIndex}", parts[3], 1), // REF
                            CreateTextCell($"C{subconDetailRowIndex}", parts[1], 1), // EQS
                            CreateTextCell($"D{subconDetailRowIndex}", parts[2], 1), // OP                            
                            CreateTextCell($"E{subconDetailRowIndex}", status, 1),     // Status
                            CreateTextCell($"F{subconDetailRowIndex}", parts[6], 1),  // Access
                            CreateTextCell($"G{subconDetailRowIndex}", parts[7], 1),  // Add
                            CreateTextCell($"H{subconDetailRowIndex}", parts[8], 1)  // Summary
                        );
                        subconSheetData.Append(subconDetailRow);
                        subconDetailRowIndex++;
                    }
                }





















                // Create individual EQS sheets
                uint sheetIdCounter = 5; // Start from 3 since main sheet is 1 and wiggets is 2
                foreach (string eqs in eqsTeam)
                {
                    var eqsSheetPart = workbookPart.AddNewPart<WorksheetPart>();

                    var eqsColumns = new OpenXMLSheet.Columns(
                        new OpenXMLSheet.Column { Min = 1, Max = 1, Width = 15, CustomWidth = true },
                        new OpenXMLSheet.Column { Min = 2, Max = 2, Width = 25, CustomWidth = true },
                        new OpenXMLSheet.Column { Min = 3, Max = 3, Width = 30, CustomWidth = true },
                        new OpenXMLSheet.Column { Min = 4, Max = 4, Width = 30, CustomWidth = true },
                        new OpenXMLSheet.Column { Min = 5, Max = 5, Width = 30, CustomWidth = true }
                    );

                    var eqsSheetData = new OpenXMLSheet.SheetData();
                    eqsSheetPart.Worksheet = new OpenXMLSheet.Worksheet(eqsColumns, eqsSheetData);

                    string sheetName = eqs.Replace(" ", "_").ToUpper();
                    string finalSheetName = sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;

                    sheets.Append(new OpenXMLSheet.Sheet
                    {
                        Id = workbookPart.GetIdOfPart(eqsSheetPart),
                        SheetId = sheetIdCounter,
                        Name = finalSheetName
                    });

                    // Create header row for individual sheet
                    var eqsHeaderRow = new OpenXMLSheet.Row() { RowIndex = 1 };
                    eqsHeaderRow.Append(
                        CreateTextCell("A1", "Type", 2),
                        CreateTextCell("B1", "Date", 2),
                        CreateTextCell("C1", "REF", 2),
                        CreateTextCell("D1", "EQS", 2),
                        CreateTextCell("E1", "OP", 2)

                    );
                    eqsSheetData.Append(eqsHeaderRow);

                    uint eqsRowIndex = 2;

                    void AppendEntries(IEnumerable<string> entries, string type)
                    {
                        foreach (var entry in entries.Where(e => e.ToLower().Contains(eqs.Split(" ").Last().ToLower())))
                        {
                            var parts = entry.Split(':', StringSplitOptions.TrimEntries);
                            if (parts.Length >= 4) // Ensure we have enough parts
                            {
                                var eqsDataRow = new OpenXMLSheet.Row() { RowIndex = eqsRowIndex };
                                eqsDataRow.Append(
                                    CreateTextCell($"A{eqsRowIndex}", type, 1),
                                    CreateTextCell($"B{eqsRowIndex}", parts[0], 1),
                                    CreateTextCell($"C{eqsRowIndex}", parts[3], 1),
                                    CreateTextCell($"D{eqsRowIndex}", parts[1], 1),
                                    CreateTextCell($"E{eqsRowIndex}", parts[2], 1)
                                );
                                eqsSheetData.Append(eqsDataRow);
                                eqsRowIndex++;
                            }
                        }
                    }





                    AppendEntries(audits, "Audit");
                    AppendEntries(post, "Post");
                    AppendEntries(wig, "Post");

                    sheetIdCounter++;
                }

                // Final save
                workbookPart.Workbook.Save();
            }

            static OpenXMLSheet.Cell CreateTextCell(string cellReference, string text, uint styleIndex)
            {
                return new OpenXMLSheet.Cell
                {
                    CellReference = cellReference,
                    DataType = OpenXMLSheet.CellValues.String,
                    CellValue = new OpenXMLSheet.CellValue(text),
                    StyleIndex = styleIndex
                };
            }

            static OpenXMLSheet.Cell CreateIntegerCell(string cellReference, int number, uint styleIndex)
            {
                return new OpenXMLSheet.Cell
                {
                    CellReference = cellReference,
                    DataType = OpenXMLSheet.CellValues.Number, // Optional for numbers
                    CellValue = new OpenXMLSheet.CellValue(number.ToString()),
                    StyleIndex = styleIndex
                };
            }

            static WorkbookStylesPart AddStyles(WorkbookPart workbookPart)
            {
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new OpenXMLSheet.Stylesheet();

                // Fonts
                var fonts = new OpenXMLSheet.Fonts();
                fonts.AppendChild(new OpenXMLSheet.Font()); // Index 0 = default
                fonts.AppendChild(new OpenXMLSheet.Font( // Index 1 = bold
                    new OpenXMLSheet.Bold()));

                // Fills (unused, but required)
                var fills = new OpenXMLSheet.Fills();
                fills.AppendChild(new OpenXMLSheet.Fill(new OpenXMLSheet.PatternFill() { PatternType = OpenXMLSheet.PatternValues.None }));
                fills.AppendChild(new OpenXMLSheet.Fill(new OpenXMLSheet.PatternFill() { PatternType = OpenXMLSheet.PatternValues.Gray125 }));

                // Borders (unused, but required)
                var borders = new OpenXMLSheet.Borders();
                borders.AppendChild(new OpenXMLSheet.Border());

                // CellFormats
                var cellFormats = new OpenXMLSheet.CellFormats();

                // Index 0: default
                cellFormats.AppendChild(new OpenXMLSheet.CellFormat());

                // Index 1: center alignment
                cellFormats.AppendChild(new OpenXMLSheet.CellFormat
                {
                    Alignment = new OpenXMLSheet.Alignment { Horizontal = OpenXMLSheet.HorizontalAlignmentValues.Center },
                    ApplyAlignment = true
                });

                // Index 2: bold + center
                cellFormats.AppendChild(new OpenXMLSheet.CellFormat
                {
                    FontId = 1,
                    Alignment = new OpenXMLSheet.Alignment { Horizontal = OpenXMLSheet.HorizontalAlignmentValues.Center },
                    ApplyFont = true,
                    ApplyAlignment = true
                });

                stylesPart.Stylesheet.Append(fonts, fills, borders, cellFormats);
                stylesPart.Stylesheet.Save();

                return stylesPart;
            }

        }

        public static string[] InputAuditDate()
        {
            string year;
            string month;

            while (true)
            {
                string[] validYears = { "2025", "2026", "2027", "2028", "2029", "2030" };

                Console.WriteLine("Please enter the year we are collecting data for, e.g. 2025:");
                string inputYear = Console.ReadLine();

                if (validYears.Any(option => option.Equals(inputYear, StringComparison.OrdinalIgnoreCase)))
                {
                    year = inputYear;

                    while (true)
                    {
                        string[] validMonths = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12" };

                        Console.WriteLine("Please enter the month we are collecting data for, e.g. for May enter 5:");
                        string inputMonth = Console.ReadLine();

                        if (validMonths.Any(option => option.Equals(inputMonth, StringComparison.OrdinalIgnoreCase)))
                        {
                            month = inputMonth;
                            Console.WriteLine($"You entered year: {year} and month: {month}");
                            // Exit both loops
                            return [year, month];
                        }
                        else
                        {
                            Console.WriteLine("Invalid month. Please try again.");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Invalid year. Please try again.");
                }
            }
        }
    }
}
