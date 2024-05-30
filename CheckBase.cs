using System.Collections.Generic;
using System.Linq;
using Models.Forms;
using OfficeOpenXml;
using System.Globalization;
using System.IO;
using System;
using System.ComponentModel;

namespace Client_App.Commands.SyncCommands.CheckForm;

public abstract class CheckBase
{
    private protected static Dictionary<string, double> D = new();

    private protected static List<Dictionary<string, string>> OKSM = new();

    private protected static List<Dictionary<string, string>> R = new();

    private protected static bool DB_Ignore = true;

    #region OverdueCalculations
    //вычисление нарушения сроков предоставления отчётов с учетом праздников

    //праздники из года в год
    private protected static List<DateTime> holidays_generic = new()
    {
        new DateTime(1,1,1), //1 января
        new DateTime(1,1,2), //2 января
        new DateTime(1,1,3), //3 января
        new DateTime(1,1,4), //4 января
        new DateTime(1,1,5), //5 января
        new DateTime(1,1,6), //6 января
        new DateTime(1,1,7), //7 января
        new DateTime(1,1,8), //8 января
        new DateTime(1,2,23), //23 февраля
        new DateTime(1,3,8), //8 марта
        new DateTime(1,5,1), //1 мая
        new DateTime(1,5,9), //9 мая
        new DateTime(1,6,12), //12 июня
        new DateTime(1,11,4), //4 ноября
    };

    //праздники конкретных годов, берутся из файла ./Spravochniki/Holidays.xlsx
    private protected static List<DateTime> holidays_specific = new()
    {

    };

    //собственно функция импорта праздничных дат из справочника
    private protected static void Holidays_Populate_From_File(string file_address)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        if (!File.Exists(file_address)) return;
        FileInfo excel_import_file = new(file_address);
        var xls = new ExcelPackage(excel_import_file);
        var wrksht1 = xls.Workbook.Worksheets["Лист1"];
        string value_date;
        DateTime fix_date;
        var i = 1;
        holidays_specific.Clear();
        while (wrksht1.Cells[i, 1].Text != string.Empty)
        {
            value_date = wrksht1.Cells[i, 1].Text;
            if (DateTime.TryParse(value_date, out fix_date))
            {
                holidays_specific.Add(fix_date);
            }
            i++;
        }
    }

    //расчет кол-ва рабочих дней между двумя датами
    //strict_order - ожидать даты в правильном порядке
    private protected static int Workdays_Between_Dates(DateTime date1, DateTime date2, bool strict_order = true)
    {
        int result = 0;
        DateTime date_min;
        DateTime date_max;
        if (strict_order)
        {
            date_min = date1;
            date_max = date2;
        }
        else
        {
            date_min = (date1 < date2 ? date1 : date2);
            date_max = (date_min == date2 ? date1 : date2);
        }
        if (date_min > date_max)
        {
            return int.MaxValue;
        }
        for (var day = date_min.Date; day < date_max.Date; day = day.AddDays(1))
        {
            if (!(day.DayOfWeek == DayOfWeek.Saturday
                || day.DayOfWeek == DayOfWeek.Sunday
                || holidays_specific.Any(x => Equals(x.Date, day.Date))
                || holidays_generic.Any(x => Equals(x.Date.Month, day.Date.Month) && Equals(x.Date.Day, day.Date.Day))
                ))
            {
                result++;
            }
        }
        return result;
    }

    #endregion

    #region CheckNotePresence

    private protected static bool CheckNotePresence(List<Form> forms, List<Note> notes, int line, byte graphNumber)
    {
        var valid = false;
        foreach (var note in notes)
        {
            if (note.RowNumber_DB == null || forms[line].ReportId == null) continue;
            var noteRowsReal = note.RowNumber_DB.Replace(" ", string.Empty);
            List<int> noteRowsFinalInt = new();
            List<string> noteRowsRealStr = new(noteRowsReal.Split(','));
            foreach (var noteRowCluster in noteRowsRealStr)
            {
                if (noteRowCluster.Contains('-'))
                {
                    var noteRowBounds = noteRowCluster.Split('-');
                    if (noteRowBounds.Length != 2
                        || !int.TryParse(noteRowBounds[0], out var noteRowsBegin)
                        || !int.TryParse(noteRowBounds[1], out var noteRowsEnd)) continue;
                    for (var i = noteRowsBegin; i <= noteRowsEnd; i++)
                    {
                        noteRowsFinalInt.Add(i);
                    }
                }
                else
                {
                    if (int.TryParse(noteRowCluster, out var noteRowClusterInt))
                    {
                        noteRowsFinalInt.Add(noteRowClusterInt);
                    }
                }
            }
            if (noteRowsFinalInt.Any(noteRowNumber =>
                    noteRowNumber == line + 1
                    && note.GraphNumber_DB == graphNumber.ToString()
                    && !string.IsNullOrWhiteSpace(note.Comment_DB)))
            {
                valid = true;
                break;
            }
        }
        return valid;
    }

    #endregion

    #region DFromFile

    private protected static void D_Populate_From_File(string file_address)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        if (!File.Exists(file_address)) return;
        FileInfo excel_import_file = new(file_address);
        var xls = new ExcelPackage(excel_import_file);
        var wrksht1 = xls.Workbook.Worksheets["Лист1"];
        var i = 2;
        string name_1, name_2, name_base, name_real;
        name_base = "аврорий";
        string value_base;
        double value_real;
        D.Clear();
        while (wrksht1.Cells[i, 1].Text != string.Empty)
        {
            name_1 = wrksht1.Cells[i, 2].Text;
            name_2 = wrksht1.Cells[i, 3].Text;
            if (name_1 != string.Empty)
            {
                name_base = name_1.ToLower();
            }
            if (name_2.Contains('-'))
            {
                if (name_2.Contains('+'))
                {
                    name_real = name_base + name_2[name_2.IndexOf('-')..name_2.IndexOf('+')];
                }
                else
                {
                    name_real = name_base + name_2[name_2.IndexOf('-')..];
                }
                value_base = wrksht1.Cells[i, 4].Text;
                if (value_base.Contains("Неограниченно"))
                {
                    value_real = double.MaxValue;
                }
                else
                {
                    value_real = 1e12 * double.Parse(value_base[..6].Replace(" ", ""), NumberStyles.Float);
                }
                D[name_real] = value_real;
                if (name_real.Contains("йод"))
                {
                    D[name_real.Replace('й', 'и')] = value_real;
                }
                else if (name_real.Contains("иод"))
                {
                    D[name_real.Replace('и', 'й')] = value_real;
                }
            }
            i++;
        }
    }

    #endregion

    #region OKSMFromFile

    private protected static void OKSM_Populate_From_File(string fileAddress)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        if (!File.Exists(fileAddress)) return;
        FileInfo excel_import_file = new(fileAddress);
        var xls = new ExcelPackage(excel_import_file);
        var wrksht1 = xls.Workbook.Worksheets["Лист1"];
        var i = 8;
        OKSM.Clear();
        while (wrksht1.Cells[i, 1].Text != string.Empty)
        {
            OKSM.Add(new Dictionary<string, string>
            {
                {"kod", wrksht1.Cells[i, 2].Text},
                {"shortname", wrksht1.Cells[i, 3].Text},
                {"longname", wrksht1.Cells[i, 4].Text},
                {"alpha2", wrksht1.Cells[i, 5].Text},
                {"alpha3", wrksht1.Cells[i, 6].Text}
            });
            i++;
        }
    }

    #endregion

    #region RFromFile

    private protected static void R_Populate_From_File(string filePath)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        if (!File.Exists(filePath)) return;
        FileInfo excelImportFile = new(filePath);
        var xls = new ExcelPackage(excelImportFile);
        var worksheet = xls.Workbook.Worksheets["Лист1"];
        var i = 2;
        R.Clear();
        while (worksheet.Cells[i, 1].Text != string.Empty)
        {
            R.Add(new Dictionary<string, string>
            {
                {"name", worksheet.Cells[i, 1].Text},
                {"value", worksheet.Cells[i, 5].Text},
                {"unit", worksheet.Cells[i, 6].Text},
                {"code", worksheet.Cells[i, 8].Text}
            });
            i++;
        }
    }

    #endregion
}