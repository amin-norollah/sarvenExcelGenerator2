using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace sarvenExcelGenerator
{
    public partial class MainWindow : Window
    {
        bool isPersian = false;

        public MainWindow()
        {
            InitializeComponent();

            MonthPicker.SelectedDate = DateTime.Now;
        }

        private void OnGenerateExcelClick(object sender, RoutedEventArgs e)
        {
            var selectedDate = MonthPicker.SelectedDate ?? DateTime.Now;
            var selectedLanguage = ((ComboBoxItem)LanguageSelector.SelectedItem).Content.ToString();

            isPersian = selectedLanguage == "Persian";
            GenerateExcel(selectedDate, isPersian);
        }

        private void GenerateExcel(DateTime selectedDate, bool isPersian)
        {
            try
            {
                System.Globalization.Calendar calendar;
                if (isPersian)
                {
                    calendar = new PersianCalendar();
                }
                else
                {
                    calendar = new GregorianCalendar();
                }

                var culture = isPersian ? new CultureInfo("fa-IR") : new CultureInfo("en-US");
                int year = calendar.GetYear(selectedDate);
                int month = calendar.GetMonth(selectedDate);

                string monthName = culture.DateTimeFormat.GetMonthName(month);

                // Parse user-provided holiday dates
                var holidayDates = ParseHolidayDates(HolidayInput.Text);

                var entries = new List<ScheduleEntry>();

                int daysInMonth = isPersian
                ? GetDaysInPersianMonth(year, month)
                : DateTime.DaysInMonth(year, month);
                for (int day = 1; day <= daysInMonth; day++)
                {
                    var date = calendar.ToDateTime(year, month, day, 0, 0, 0, 0);

                    bool isHoliday = holidayDates.Contains(day);

                    entries.Add(new ScheduleEntry
                    {
                        DayOfWeek = date.ToString("dddd", culture),
                        Date = date.ToString("yyyy/MM/dd", culture),
                        StartHour = "",
                        EndHour = "",
                        Difference = "",
                        Description = "",
                        IsHoliday = isHoliday || date.DayOfWeek == DayOfWeek.Thursday || date.DayOfWeek == DayOfWeek.Friday
                    });
                }

                SaveExcel(entries, isPersian ? "کارمند" : "Employee", monthName, year);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Parse the holiday dates from user input
        private List<int> ParseHolidayDates(string input)
        {
            return input.Split(',')
                        .Where(date => int.TryParse(date.Trim(), out _))
                        .Select(date => int.Parse(date.Trim()))
                        .ToList();
        }

        private int GetDaysInPersianMonth(int year, int month)
        {
            var persianCalendar = new PersianCalendar();
            DateTime firstDayOfMonth = persianCalendar.ToDateTime(year, month, 1, 0, 0, 0, 0);
            DateTime firstDayOfNextMonth;

            if (month == 12)
            {
                firstDayOfNextMonth = persianCalendar.ToDateTime(year + 1, 1, 1, 0, 0, 0, 0);
            }
            else
            {
                firstDayOfNextMonth = persianCalendar.ToDateTime(year, month + 1, 1, 0, 0, 0, 0);
            }

            return (int)(firstDayOfNextMonth - firstDayOfMonth).TotalDays;
        }

        private void SaveExcel(List<ScheduleEntry> entries, string employeeName, string monthName, int year)
        {
            string filePath = "";// Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{year}-{monthName}.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Schedule");

                // Add Month and Year in the First Row
                worksheet.Row(1).Height = 36;
                worksheet.Cell(1, 1).Value = isPersian
                    ? $"فرم گزارش کار و ثبت ساعت کاری {monthName} ({year})\n نام و نام خانوادگی: "
                    : $"Work Report Form for {monthName} ({year})\n Name and Family: ";
                worksheet.Range(1, 1, 1, 6).Merge();
                worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell(1, 1).Style.Alignment.WrapText = true;
                worksheet.Cell(1, 1).Style.Font.Bold = true;

                // Header Row
                worksheet.Row(2).Height = 30;
                worksheet.Cell(2, 1).Value = isPersian ? "ایام هفته" : "Day of Week";
                worksheet.Cell(2, 2).Value = isPersian ? "تاریخ" : "Date";
                worksheet.Cell(2, 3).Value = isPersian ? "ساعت ورود" : "Start Hour";
                worksheet.Cell(2, 4).Value = isPersian ? "ساعت خروج" : "End Hour";
                worksheet.Cell(2, 5).Value = isPersian ? "کارکرد" : "Difference";
                worksheet.Cell(2, 6).Value = isPersian ? "شرح" : "Description";
                worksheet.Row(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Row(2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                worksheet.Row(2).Style.Font.Bold = true;
                for (int col = 1; col <= 6; col++)
                {
                    worksheet.Cell(2, col).Style.Fill.BackgroundColor = XLColor.LightGray;
                }

                // Set column widths
                worksheet.Column(1).Width = 12; // Day of Week
                worksheet.Column(2).Width = 12; // Date
                worksheet.Column(5).Width = 12; // Difference
                worksheet.Column(6).Width = 60; // Description

                // Data Rows
                int row = 3;
                foreach (var entry in entries)
                {
                    worksheet.Row(row).Height = 22;
                    worksheet.Row(row).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Row(row).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    worksheet.Cell(row, 1).Value = entry.DayOfWeek;
                    worksheet.Cell(row, 2).Value = entry.Date;
                    worksheet.Cell(row, 3).Value = entry.StartHour;
                    worksheet.Cell(row, 4).Value = entry.EndHour;
                    worksheet.Cell(row, 5).Value = entry.Difference;
                    worksheet.Cell(row, 6).Value = entry.Description;

                    if (entry.IsHoliday)
                    {
                        for (int col = 1; col <= 6; col++)
                        {
                            worksheet.Cell(row, col).Style.Fill.BackgroundColor = XLColor.FromHtml("#ffefa8");
                        }
                    }
                    worksheet.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#ffe054");

                    // Add border to all cells
                    for (int col = 1; col <= 6; col++)
                    {
                        worksheet.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    row++;
                }

                // Calculate sum for 'Difference' column
                worksheet.Cell(row, 1).Value = isPersian ? "جمع کل" : "Total";
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range(row, 1, row, 4).Merge(); // Merge first 4 cells in the last row
                worksheet.Cell(row, 5).FormulaA1 = $"SUM(E3:E{row - 1})";
                worksheet.Cell(row, 5).Style.DateFormat.Format = "h:mm:ss";
                worksheet.Cell(row, 5).Style.Font.Bold = true;

                // Style the total row
                worksheet.Row(row).Height = 30;
                worksheet.Row(row).Style.Font.Bold = true;
                worksheet.Cell(row, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Row(row).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Row(row).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                for (int col = 1; col <= 6; col++)
                {
                    worksheet.Cell(row, col).Style.Fill.BackgroundColor = XLColor.FromHtml("#ffe054");
                    worksheet.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // Apply border to header row
                for (int col = 1; col <= 6; col++)
                {
                    worksheet.Cell(2, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                // Save to File
                // Create a SaveFileDialog instance
                SaveFileDialog saveFileDialog = new SaveFileDialog();

                // Set default file extension and filter
                saveFileDialog.FileName = $"{year}-{monthName}.xlsx";
                saveFileDialog.DefaultExt = ".xlsx";
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"; // Filter for Excel files


                // Show the SaveFileDialog and check if the user selected a file
                if (saveFileDialog.ShowDialog() == true)
                {
                    // Get the file path chosen by the user
                    filePath = saveFileDialog.FileName;

                    // You can now save your file to the selected location
                    try
                    {
                        workbook.SaveAs(filePath);
                        MessageBox.Show("Excel file generated successfully! " + filePath, "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }

            }
        }

        private class ScheduleEntry
        {
            public string DayOfWeek { get; set; }
            public string Date { get; set; }
            public string StartHour { get; set; }
            public string EndHour { get; set; }
            public string Difference { get; set; }
            public string Description { get; set; }
            public bool IsHoliday { get; set; }
        }
    }
}