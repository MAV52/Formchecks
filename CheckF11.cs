using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Models.CheckForm;
using Models.Collections;
using Models.Forms;
using Models.Forms.Form1;
using OfficeOpenXml;

namespace Client_App.Commands.SyncCommands.CheckForm;

public class CheckF11
{
    private static readonly string[] OperationCode_DB_Valids =
    {
        "1","10","11","12","15","17","18","21",
        "22","25","27","28","29","31","32","35",
        "37","38","39","41","42","43","48","53",
        "54","58","61","62","63","64","65","66",
        "67","68","71","72","73","74","75","81",
        "82","83","84","85","86","87","88","97",
        "98","99"
    };

    private static readonly string[] Radionuclids_DB_Valids =
    {
        "плутоний","уран-233","уран-235","нептуний-237","америций-241","америций-243","калифорний-252","торий","литий-6","тритий"
    };

    private static readonly Dictionary<string, string> GraphsList = new()
    {
        { "NumberInOrder_DB", "01 - № п/п" },
        { "OperationCode_DB", "02 - Код операции" },
        { "OperationDate_DB", "03 - Дата операции" },
        { "PassportNumber_DB", "04 - Номер паспорта (сертификата)" },
        { "Type_DB", "05 - Тип ЗРИ" },
        { "Radionuclids_DB", "06 - Радионуклиды" },
        { "FactoryNumber_DB", "07 - Номер" },
        { "Quantity_DB", "08 - Количество, шт." },
        { "Activity_DB", "09 - Суммарная активность, Бк" },
        { "CreatorOKPO_DB", "10 - код ОКПО изготовителя" },
        { "CreationDate_DB", "11 - Дата выпуска" },
        { "Category_DB", "12 - Категория ЗРИ" },
        { "SignedServicePeriod_DB", "13 - НСС, месяцев" },
        { "PropertyCode_DB", "14 - Код формы собственности" },
        { "Owner_DB", "15 - Код ОКПО правообладателя" },
        { "DocumentVid_DB", "16 - Вид документа" },
        { "DocumentNumber_DB", "17 - Номер документа" },
        { "DocumentDate_DB", "18 - Дата документа" },
        { "ProviderOrRecieverOKPO_DB", "19 - Код ОКПО поставщика или получателя" },
        { "TransporterOKPO_DB", "20 - Код ОКПО перевозчика" },
        { "PackName_DB", "21 - Наименование прибора, УКТ, упаковки" },
        { "PackType_DB", "22 - Тип прибора, УКТ, упаковки" },
        { "PackNumber_DB", "23 - Номер прибора, УКТ, упаковки" },
    };

    public static string[] Type_DB_Valids =
    {

    };

    private static List<Dictionary<string, string>> OKSM = new()
    {
    };

    private static Dictionary<string, double> D = new()
    {
    };

    private static bool DB_Ignore = true;

    private static bool ZRI_Ignore = true;

    private static bool MZA_Ignore = true;

    #region DFromFile

    private static void D_Populate_From_File(string file_address)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

    private static void OKSM_Populate_From_File(string fileAddress)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

    #region CheckTotal

    public static List<CheckError> Check_Total(Reports reps, Report rep)
    {
        var currentFormLine = 0;
        List<CheckError> errorList = new();
        if (OKSM.Count == 0)
        {
#if DEBUG
            OKSM_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\..\\")), "data", "Spravochniki", "oksm.xlsx"));
#else
            OKSM_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"oksm.xlsx"));
#endif
        }
        if (D.Count == 0)
        {
#if DEBUG
            D_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\..\\")), "data", "Spravochniki", "D.xlsx"));
#else
            D_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"D.xlsx"));
#endif
        }
        foreach (var key in rep.Rows11)
        {
            var form = (Form11)key;
            var formsList = rep.Rows11.ToList<Form11>();
            var notes = rep.Notes.ToList<Note>();
            var forms10 = reps.Master_DB.Rows10.ToList<Form10>();
            errorList.AddRange(Check_001(formsList, currentFormLine));
            errorList.AddRange(Check_002(formsList, currentFormLine));
            errorList.AddRange(Check_003(formsList, currentFormLine));
            errorList.AddRange(Check_004(formsList, currentFormLine));
            errorList.AddRange(Check_005(formsList, currentFormLine));
            errorList.AddRange(Check_006(formsList, currentFormLine));
            errorList.AddRange(Check_007(formsList, notes, currentFormLine));
            errorList.AddRange(Check_008(formsList, currentFormLine));
            errorList.AddRange(Check_009(formsList, currentFormLine));
            errorList.AddRange(Check_010(formsList, currentFormLine));
            errorList.AddRange(Check_011(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_012(formsList, currentFormLine));
            errorList.AddRange(Check_013(formsList, currentFormLine));
            errorList.AddRange(Check_014(formsList, currentFormLine));
            errorList.AddRange(Check_015(formsList, currentFormLine));
            errorList.AddRange(Check_016(formsList, currentFormLine));
            errorList.AddRange(Check_017(formsList, rep, currentFormLine));
            errorList.AddRange(Check_018(formsList, rep, currentFormLine));
            errorList.AddRange(Check_019(formsList, currentFormLine));
            errorList.AddRange(Check_020(formsList, currentFormLine));
            errorList.AddRange(Check_021(formsList, currentFormLine));
            errorList.AddRange(Check_022(formsList, currentFormLine));
            errorList.AddRange(Check_023(formsList, currentFormLine));
            errorList.AddRange(Check_024(formsList, currentFormLine));
            errorList.AddRange(Check_025(formsList, currentFormLine));
            errorList.AddRange(Check_026(formsList, currentFormLine));
            errorList.AddRange(Check_027(formsList, currentFormLine));
            errorList.AddRange(Check_028(formsList, currentFormLine));
            errorList.AddRange(Check_029(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_030(formsList, currentFormLine));
            errorList.AddRange(Check_031(formsList, notes, currentFormLine));
            errorList.AddRange(Check_032(formsList, currentFormLine));
            errorList.AddRange(Check_033(formsList, currentFormLine));
            errorList.AddRange(Check_034(formsList, currentFormLine));
            errorList.AddRange(Check_035(formsList, currentFormLine));
            errorList.AddRange(Check_036(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_037(formsList, currentFormLine));
            errorList.AddRange(Check_038(formsList, notes, currentFormLine));
            errorList.AddRange(Check_039(formsList, notes, currentFormLine));
            errorList.AddRange(Check_040(formsList, notes, currentFormLine));
            errorList.AddRange(Check_041(formsList, currentFormLine));
            errorList.AddRange(Check_042(formsList, currentFormLine));
            errorList.AddRange(Check_043(formsList, currentFormLine));
            errorList.AddRange(Check_044(formsList, rep, currentFormLine));
            errorList.AddRange(Check_045(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_046(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_047(formsList, currentFormLine));
            errorList.AddRange(Check_048(formsList, notes, currentFormLine));
            errorList.AddRange(Check_049(formsList, currentFormLine));
            errorList.AddRange(Check_050(formsList, currentFormLine));
            errorList.AddRange(Check_051(formsList, currentFormLine));
            errorList.AddRange(Check_052(formsList, currentFormLine));
            errorList.AddRange(Check_053(formsList, currentFormLine));
            errorList.AddRange(Check_054(formsList, currentFormLine));
            errorList.AddRange(Check_055(formsList, currentFormLine));

            currentFormLine++;
        }

        var index = 0;
        foreach (var error in errorList)
        {
            if (GraphsList.TryGetValue(error.Column, out var columnFrontName))
            {
                error.Column = columnFrontName;
            }
            index++;
            error.Index = index;
        }
        return errorList;
    }

    #endregion

    #region Checks

    #region Check001
    
    private static List<CheckError> Check_001(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        return result;
    }

    #endregion

    #region Check002
    
    private static List<CheckError> Check_002(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].Id < 1)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Id",
                Value = forms[line].Id.ToString(),
                Message = "-"
            });
        }
        return result;
    }

    #endregion

    #region Check003
    
    private static List<CheckError> Check_003(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var valid = operationCode != null && OperationCode_DB_Valids.Contains(operationCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check004
    
    private static List<CheckError> Check_004(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = forms[line].OperationCode_DB,
                Message = "Сведения, представленные в инвентаризации, не соответствуют СНК."
            });
        }
        return result;
    }

    #endregion

    #region Check005
    
    //Для кодов операции 12, 42, в радионуклидах должен быть указан хоть один из списка (колонка 6)
    private static List<CheckError> Check_005(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "12", "42" };
        var radionuclids = forms[line].Radionuclids_DB;
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = Radionuclids_DB_Valids.Any(nuclid => 
            radionuclids?.Contains(nuclid, StringComparison.CurrentCultureIgnoreCase) == true);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = forms[line].Radionuclids_DB,
                Message = "В графе 6 не представлены сведения о радионуклидах, которые могут быть отнесены к ЯМ. Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check006
    
    private static List<CheckError> Check_006(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "15" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "-",
                Value = "-",
                Message = "В предыдущих отчетах не найдена строка об осуществлении операции переработки РАО в виде ОЗИИИ. Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check007
    
    //Наличие примечания для кодов операции 29, 29, 97-99
    private static List<CheckError> Check_007(List<Form11> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "29", "39", "97", "98", "99" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
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
                    && note.GraphNumber_DB == 2.ToString() 
                    && !string.IsNullOrWhiteSpace(note.Comment_DB)))
            {
                valid = true;
                break;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = "-",
                Message = "Необходимо дать пояснение об осуществленной операции."
            });
        }
        return result;
    }

    #endregion

    #region Check008
    
    private static List<CheckError> Check_008(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "21", "22", "25", "27", "28", "29", "41", "42", "43", "46", "53", "54", "61", "62", "65", "66", "67", "68", "71", "72", "81", "82", "83", "84", "88", "98" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "-",
                Value = "-",
                Message = "Учетной единицы с такими параметрами нет в организации. Проверьте правильность указываемых сведений для ЗРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check009
    
    private static List<CheckError> Check_009(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "37" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = forms[line].OperationCode_DB,
                Message = "В отчетах не найдена строка об осуществлении передачи учетной единицы. Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check010
    
    private static List<CheckError> Check_010(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "41" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "-",
                Value = "-",
                Message = "Заполните форму 1.5."
            });
        }
        return result;
    }

    #endregion

    #region Check011
    
    //Для кода операции 54, ОКПО получателя/перевозчика не должен совпадать с ОКПО организации
    private static List<CheckError> Check_011(List<Form11> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "54" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        var valid = providerOrRecieverOKPO != repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOKPO,
                Message = "В графе 19 необходимо указать ОКПО подрядной организации."
            });
        }
        return result;
    }

    #endregion

    #region Check012
    
    private static List<CheckError> Check_012(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "58" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = forms[line].OperationCode_DB,
                Message = "В отчетах не найдены сведения о снятии с учета учетной единицы для разукомплектования. Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check013
    
    private static List<CheckError> Check_013(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "62" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "-",
                Value = forms[line].OperationCode_DB,
                Message = "В отчетах не найдены сведения о вывозе учетной единицы. Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check014
    
    private static List<CheckError> Check_014(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "65" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "-",
                Value = "-",
                Message = "Заполните форму 1.3."
            });
        }
        return result;
    }

    #endregion

    #region Check015
    
    //Код ОКПО состоит из 8/14 чисел для определенных кодов операции (колонка 10)
    private static List<CheckError> Check_015(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "81", "88" };
        var operationCode = forms[line].OperationCode_DB;
        var сreatorOKPO = forms[line].CreatorOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(сreatorOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = forms[line].CreatorOKPO_DB,
                Message = "Код используется для предоставления сведений о ЗРИ, произведенных в Российской Федерации."
            });
        }
        return result;
    }

    #endregion

    #region Check016
    
    //Коды операции 83-86 используются для зарубежных стран 
    private static List<CheckError> Check_016(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "83", "84", "85", "86" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = OKSM.All(oksmEntry => oksmEntry["shortname"] != forms[line].CreatorOKPO_DB) 
                    || forms[line].CreatorOKPO_DB.ToLower() is "россия";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = forms[line].CreatorOKPO_DB,
                Message = "Код используется для предоставления сведений о ЗРИ, произведенных за пределами Российской Федерации."
            });
        }
        return result;
    }

    #endregion

    #region Check017
    
    //Дата документа входит в отчетный период с учетом срока подачи отчета в днях (колонка 3)
    private static List<CheckError> Check_017(List<Form11> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] nonApplicableOperationCodes = { "10" };
        var opCode = forms[line].OperationCode_DB;
        if (nonApplicableOperationCodes.Contains(opCode)) return result;
        var valid = forms[line].OperationDate_DB != null;
        var pEnd = DateTime.MinValue;
        var pMid = DateTime.MinValue;
        if (valid && rep is { StartPeriod_DB: not null, EndPeriod_DB: not null })
        {
            valid = DateTime.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateTime.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateTime.TryParse(forms[line].OperationDate_DB!, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = forms[line].OperationDate_DB,
                Message = "Дата операции не входит в отчетный период."
            });
        }
        else
        {
            string[] operationCodeWithDeadline1 = { "71" };
            string[] operationCodeWithDeadline5 = { "73", "74", "75" };
            string[] operationCodeWithDeadline10 =
            {
                "11", "12", "15", "17", "18", "21", "22", "25", "27", "28", "29", "31", "32", "35", "37", "38", "39",
                "41", "42", "43", "46", "47", "48", "53", "54", "58", "61", "62", "63", "64", "65", "66", "67", "68",
                "72", "81", "82", "83", "84", "85", "86", "87", "88", "97", "98", "99"
            };
            if (operationCodeWithDeadline10.Contains(opCode) && (pEnd - pMid).Days > 10)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_11",
                    Row = (line + 1).ToString(),
                    Column = "OperationDate_DB",
                    Value = forms[line].OperationDate_DB,
                    Message = "Дата окончания отчетного периода превышает дату операции более чем на 10 дней."
                });
            }
            else if (operationCodeWithDeadline5.Contains(opCode) && (pEnd - pMid).Days > 5)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_11",
                    Row = (line + 1).ToString(),
                    Column = "OperationDate_DB",
                    Value = forms[line].OperationDate_DB,
                    Message = "Дата окончания отчетного периода превышает дату операции более чем на 5 дней."
                });
            }
            else if (operationCodeWithDeadline1.Contains(opCode) && (pEnd - pMid).Days > 1)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_11",
                    Row = (line + 1).ToString(),
                    Column = "OperationDate_DB",
                    Value = forms[line].OperationDate_DB,
                    Message = "Дата окончания отчетного периода превышает дату операции более чем на 1 день."
                });
            }
        }
        return result;
    }

    #endregion

    #region Check018
    
    //Дата документа входит в отчетный период при коде операции 10 (колонка 3) + 10 дней
    private static List<CheckError> Check_018(List<Form11> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10" };
        var documentDate = forms[line].DocumentDate_DB;
        var operationDate = forms[line].OperationDate_DB;

        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var pEnd = DateTime.MinValue;
        var pMid = DateTime.MinValue;
        var valid = DateTime.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateTime.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateTime.TryParse(documentDate, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата акта инвентаризации не входит в отчетный период."
            });
        }
        else if ((pEnd - pMid).Days > 10)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = operationDate,
                Message = "Дата окончания отчетного периода превышает дату операции более чем на 10 дней."
            });
        }
        return result;
    }

    #endregion

    #region Check019
    
    //Номер паспорта не пустая строка (колонка 4)
    private static List<CheckError> Check_019(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var valid = !string.IsNullOrWhiteSpace(forms[line].PassportNumber_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "PassportNumber_DB",
                Value = forms[line].PassportNumber_DB,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check020
    
    private static List<CheckError> Check_020(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (ZRI_Ignore) return result;
        var valid = !string.IsNullOrEmpty(forms[line].Type_DB);
        if (valid)
        {
            valid = Type_DB_Valids.Contains(forms[line].Type_DB);
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Type_DB",
                Value = forms[line].Type_DB,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check021
    
    private static List<CheckError> Check_021(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (ZRI_Ignore) return result;
        var valid = !string.IsNullOrEmpty(forms[line].Type_DB)
                    && !string.IsNullOrEmpty(forms[line].Radionuclids_DB);
        if (!Type_DB_Valids.Contains(forms[line].PackType_DB)) return result;
        if (valid)
        {
            List<string> radionuclidsValidFull = new();
            List<string> radionuclidsFactFull = new(forms[line].Radionuclids_DB!.Replace(" ", string.Empty).ToLower().Split(";"));
            foreach (var radionuclidFact in radionuclidsFactFull)
            {
                if (!radionuclidsValidFull.Contains(radionuclidFact))
                {
                    valid = false;
                    break;
                }
                else
                {
                    radionuclidsValidFull.Remove(radionuclidFact);
                }
            }
            valid = radionuclidsValidFull.Count == 0;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = forms[line].Radionuclids_DB,
                Message = "Перечень радионуклидов не соответствует справочным данным."
            });
        }
        return result;
    }

    #endregion

    #region Check022
    
    //У радионуклидов в качестве разделителя не могут выступать ',' и '+'
    private static List<CheckError> Check_022(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        char[] excludedDelimiters = { ',', '+' };
        var radionuclids = forms[line].Radionuclids_DB;
        if (string.IsNullOrEmpty(radionuclids)) return result;
        var valid = radionuclids.IndexOfAny(excludedDelimiters) == -1;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = forms[line].Radionuclids_DB,
                Message = "Формат ввода данных не соответствует приказу"
            });
        }
        return result;
    }

    #endregion

    #region Check023
    
    //Разделение через ";", если количество > 1.
    private static List<CheckError> Check_023(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var factoryNum = forms[line].FactoryNumber_DB;
        var quantity = forms[line].Quantity_DB;
        if (factoryNum == null) return result;
        var valid = !string.IsNullOrWhiteSpace(factoryNum)
                    && !factoryNum.Contains(',');
        if (quantity > 1 && !factoryNum.Contains(';')
            || quantity == 1 && factoryNum.Contains(';'))
        {
            valid = false;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "FactoryNumber_DB",
                Value = factoryNum,
                Message = "Формат ввода данных не соответствует приказу"
            });
        }
        return result;
    }

    #endregion

    #region Check024
    
    private static List<CheckError> Check_024(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var quantity = forms[line].Quantity_DB;
        var valid = quantity is > 0;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Quantity_DB",
                Value = quantity.ToString(),
                Message = "Формат ввода данных не соответствует приказу"
            });
        }
        return result;
    }

    #endregion

    #region Check025
    private static List<CheckError> Check_025(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (MZA_Ignore) return result;
        if (string.IsNullOrEmpty(forms[line].Radionuclids_DB)
            || forms[line].Activity_DB == null || forms[line].Activity_DB == string.Empty
            || forms[line].Activity_DB == "-") return result;
        var nuclidsList = forms[line].Radionuclids_DB!.ToLower().Replace(" ", string.Empty).Split(';');
        if (nuclidsList.Length != 1) return result;
        if (!Radionuclids_DB_Valids.Contains(nuclidsList[0])) return result;
        //find the minimum activity
        double activityMinimum = float.MaxValue;
        if (D.TryGetValue(nuclidsList[0], out var value))
        {
            activityMinimum = value;
        }
        var activityReal = double.Parse(forms[line].Activity_DB!.Replace(".", ","), NumberStyles.Float);
        var valid = activityReal >= activityMinimum;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = forms[line].Activity_DB,
                Message = "Активность ниже МЗА, ЗРИ не является объектом учета СГУК РВ и РАО."
            });
        }
        return result;
    }

    #endregion

    #region Check026
    
    private static List<CheckError> Check_026(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (MZA_Ignore) return result;
        if (string.IsNullOrEmpty(forms[line].Radionuclids_DB)
            || forms[line].Activity_DB == null
            || forms[line].Activity_DB == string.Empty
            || forms[line].Activity_DB == "-") return result;
        var nuclidsList = forms[line].Radionuclids_DB!.ToLower().Replace(" ", string.Empty).Split(';');
        if (nuclidsList.Length == 1) return result;
        if (nuclidsList.Any(nuclid => !Radionuclids_DB_Valids.Contains(nuclid)))
        {
            return result;
        }
        //find the minimum activity
        var activityMinimum = double.MaxValue;
        foreach (var nuclid in nuclidsList)
        {
            double activityConsidered = float.MaxValue;
            if (D.TryGetValue(nuclid, out var value))
            {
                activityConsidered = value;
            }
            if (activityConsidered < activityMinimum)
            {
                activityMinimum = activityConsidered;
            }
        }
        var activityReal = double.Parse(forms[line].Activity_DB!.Replace(".", ","), NumberStyles.Float);
        var valid = activityReal >= activityMinimum;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = forms[line].Activity_DB,
                Message = "Активность ниже МЗА, ЗРИ не является объектом учета СГУК РВ и РАО."
            });
        }
        return result;
    }

    #endregion

    #region Check027
    
    private static List<CheckError> Check_027(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        if (MZA_Ignore) return result;
        return result;
    }

    #endregion

    #region Check028

    //"Суммарная активность, Бк" положительное число, менее 10е+20
    private static List<CheckError> Check_028(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].Activity_DB;
        if (string.IsNullOrEmpty(activity) || activity == "-") return result;
        activity = forms[line].Activity_DB
            !.Replace(".", ",")
            .Replace("(", "")
            .Replace(")", "");
        if (!double.TryParse(activity,
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("en-GB"),
                out var activityReal)
            || activity.Contains('-')
            || activityReal is <= 0 or > 10e+20)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = forms[line].Activity_DB,
                Message = "Проверьте правильность предоставления сведений по суммарной активности."
            });
        }
        return result;
    }

    #endregion

    #region Check029

    //"код ОКПО изготовителя" = "ОКПО организации" для кодов операции 11, 58
    private static List<CheckError> Check_029(List<Form11> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "11" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var creatorOkpo = forms[line].CreatorOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        var valid = creatorOkpo == repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = forms[line].CreatorOKPO_DB,
                Message = "Проверьте код ОКПО организации-изготовителя."
            });
        }
        return result;
    }

    #endregion

    #region Check030

    //Соответствие формата данных "код ОКПО изготовителя" (колонка 10)
    private static List<CheckError> Check_030(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var creatorOKPO = forms[line].CreatorOKPO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = !string.IsNullOrEmpty(creatorOKPO)
                    && (okpoRegex.IsMatch(creatorOKPO)
                        || creatorOKPO is "прим."
                        || OKSM.Any(oksmEntry => oksmEntry["shortname"] == creatorOKPO));
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = creatorOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check031

    // Если в "код ОКПО изготовителя" (колонка 10) указано примечание, то проверяется наличие примечания для данной строки
    private static List<CheckError> Check_031(List<Form11> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        string[] creatorOkpoValid = { "прим.", "прим", "примечание", "примечания" };
        if (!creatorOkpoValid.Contains(forms[line].CreatorOKPO_DB.ToLower())) return result;
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
                    && note.GraphNumber_DB == 10.ToString()
                    && !string.IsNullOrWhiteSpace(note.Comment_DB)))
            {
                valid = true;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = forms[line].CreatorOKPO_DB,
                Message = "Необходимо указать в примечании наименование и адрес организации-изготовителя ЗРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check032
    
    //Дата выпуска (колонка 11) <= Дате операции (колонка 3)
    private static List<CheckError> Check_032(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var valid = DateTime.TryParse(forms[line].OperationDate_DB, out var operDate)
                    && DateTime.TryParse(forms[line].CreationDate_DB, out var createDate)
                    && createDate <= operDate;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "CreationDate_DB",
                Value = forms[line].CreationDate_DB,
                Message = "Дата выпуска не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check033
    
    //Соответствие категории ЗРИ (колонка 12)
    private static List<CheckError> Check_033(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var dbBounds = new Dictionary<short, (double, double)>
        {
            { 1, (1000, double.MaxValue) },
            { 2, (10, 1000) },
            { 3, (1, 10) },
            { 4, (0.01, 1) },
            { 5, (0, 0.01) }
        };
        var activity = forms[line].Activity_DB;
        var category = forms[line].Category_DB;
        var quantity = forms[line].Quantity_DB;
        var radionuclids = forms[line].Radionuclids_DB;
        List<double> dValueList = new();
        var nuclidsList = radionuclids!
            .ToLower()
            .Replace(" ", string.Empty)
            .Split(';');
        var valid = category != null
                    && activity != null
                    && dbBounds.ContainsKey((short)category)
                    && nuclidsList.Length > 0;
        
        if (valid)
        {
            foreach (var nuclid in nuclidsList)
            {
                if (D.TryGetValue(nuclid, out var value))
                {
                    dValueList.Add(value);
                }
            }
            if (dValueList.Count == 0)
            {
                foreach (var nuclid in nuclidsList)
                {
                    foreach (var key in D.Keys.Where(key => key.Contains(nuclid)))
                    {
                        dValueList.Add(D[key] / (quantity != null && quantity != 0 
                            ? (double)quantity 
                            : 1.0));
                        break;
                    }
                }
            }
            if (dValueList.Count == 0)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_11",
                    Row = (line + 1).ToString(),
                    Column = "Radionuclids_DB",
                    Value = radionuclids,
                    Message = "Проверьте правильность заполнения графы 6."
                });
                return result;
            }
            var dMinValue = dValueList.Min();
            var dMaxValue = dValueList.Max();
            valid = double.TryParse(activity!.Replace(".", ","), 
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands, 
                CultureInfo.CreateSpecificCulture("en-GB"),
                out var aValue);
            if (valid)
            {
                var adMinBound = dMaxValue == 0.0 
                    ? double.MaxValue 
                    : aValue / dMaxValue;
                var adMaxBound = dMinValue == 0.0 
                    ? double.MaxValue 
                    : aValue / dMinValue;
                valid = dbBounds[(short)category!].Item1 <= adMinBound
                        && dbBounds[(short)category].Item2 > adMaxBound;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Category_DB",
                Value = category.ToString(),
                Message = "Проверьте правильность указания категории ЗРИ. Проверьте активность по паспорту, должна быть указана в Бк"
            });
        }
        return result;
    }

    #endregion

    #region Check034
    
    //Дата выпуска + НСС < дата операции
    private static List<CheckError> Check_034(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var signedServicePeriod = forms[line].SignedServicePeriod_DB;
        var creationDate = forms[line].CreationDate_DB;
        var operationDate = forms[line].OperationDate_DB;

        var valid = false;
        if (signedServicePeriod != null
            && !string.IsNullOrEmpty(creationDate)
            && !string.IsNullOrEmpty(operationDate)
            && DateTime.TryParse(creationDate, out var creationDateReal)
            && DateTime.TryParse(operationDate, out var operationDateReal))
        {
            creationDateReal = creationDateReal.AddMonths((int)signedServicePeriod!);
            creationDateReal = creationDateReal.AddDays(Math.Round((double)(30 * (signedServicePeriod % 1.0))));
            valid = creationDateReal >= operationDateReal;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "SignedServicePeriod_DB",
                Value = signedServicePeriod.ToString(),
                Message = "Для ЗРИ истек НСС, следует продлить НСС либо снять с учета с одновременной постановкой на учет как РАО (при выполнении критериев отнесения к РАО)." 
                          + $"{Environment.NewLine}Проверьте, что НСС указан в месяцах." 
            });
        }
        return result;
    }

    #endregion

    #region Check035
    
    //Код формы собственности (колонка 14) от 1 до 6, 9
    private static List<CheckError> Check_035(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        byte?[] propertyCodeDBValid = { 1, 2, 3, 4, 5, 6, 9 };
        if (!propertyCodeDBValid.Contains(propertyCode))
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = propertyCode.ToString(),
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check036
    
    //8 или 14 чисел (колонка 15) если код формы собственности от 1 до 4
    private static List<CheckError> Check_036(List<Form11> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        byte?[] propertyCodeValid = { 1, 2, 3, 4 };
        string[] applicableOperationCodes = { "11", "12", "28", "38", "41", "63", "64" };
        var okpoRep = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB) 
            ? forms10[1].Okpo_DB 
            : forms10[0].Okpo_DB;
        var operationCode = forms[line].OperationCode_DB;
        var owner = forms[line].Owner_DB;
        var propertyCode = forms[line].PropertyCode_DB;

        if (applicableOperationCodes.Contains(operationCode) && okpoRep == owner)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = $"Для кода операции {operationCode}, код ОКПО правообладателя не может совпадать с ОКПО отчитывающейся организации"
            });
        }
        
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        var valid = !string.IsNullOrEmpty(owner)
                    && okpoRegex.IsMatch(owner);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check037
    
    //Правообладатель (колонка 15) из справочника ОКСМ, если код формы собственности 5
    private static List<CheckError> Check_037(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        byte?[] propertyCodeValid = { 5 };
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        var valid = OKSM.Any(oksmEntry => oksmEntry["shortname"] == owner) 
                    && !owner.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check038
    
    //Правообладатель (колонка 15) из справочника ОКСМ и наличие примечания, если код формы собственности 6
    private static List<CheckError> Check_038(List<Form11> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        byte?[] propertyCodeValid = { 6 };
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        if (OKSM.All(oksmEntry => oksmEntry["shortname"] != owner) 
            || owner.Equals("россия", StringComparison.CurrentCultureIgnoreCase))
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Необходимо указать в примечании наименование и адрес правообладателя (собственника или обладателя иного вещного права) на ЗРИ."
            });
            return result;
        }
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
                    && note.GraphNumber_DB == 15.ToString() 
                    && !string.IsNullOrWhiteSpace(note.Comment_DB)))
            {
                valid = true;
                break;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = forms[line].Owner_DB,
                Message = "Необходимо указать в примечании наименование и адрес правообладателя (собственника или обладателя иного вещного права) на ЗРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check039
    
    //Если код формы собственности 9, то должно быть примечание, а правообладатель (колонка 15) из ОКСМ или 8/14 цифр
    private static List<CheckError> Check_039(List<Form11> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        short?[] propertyCodeValid = { 9 };
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if ((OKSM.All(oksmEntry => oksmEntry["shortname"] != owner) 
            || owner.Equals("россия", StringComparison.CurrentCultureIgnoreCase)) 
            && !okpoRegex.IsMatch(owner))
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Необходимо указать в примечании наименование и адрес правообладателя (собственника или обладателя иного вещного права) на ЗРИ."
            });
            return result;
        }
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
                    && note.GraphNumber_DB == 15.ToString() 
                    && !string.IsNullOrWhiteSpace(note.Comment_DB)))
            {
                valid = true;
                break;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Необходимо указать в примечании наименование и адрес правообладателя (собственника или обладателя иного вещного права) на ЗРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check040
    
    //Вид документа от 1 до 15, 19. При коде операции 10, должен быть равен 1. При виде документа равном 19, должно быть примечание (колонка 16)
    private static List<CheckError> Check_040(List<Form11> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        byte?[] documentVidValid = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 19 };
        byte?[] documentVidValidFor10 = { 1 };
        var documentVid = forms[line].DocumentVid_DB;
        var valid = forms[line].OperationCode_DB == "10" 
            ? documentVidValidFor10.Contains(documentVid) 
            : documentVidValid.Contains(documentVid);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = documentVid.ToString(),
                Message = "Формат ввода данных не соответствует приказу."
            });
            return result;
        }
        if (documentVid is 19)
        {
            valid = false;
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
                        && note.GraphNumber_DB == 16.ToString() 
                        && !string.IsNullOrWhiteSpace(note.Comment_DB)))
                {
                    valid = true;
                    break;
                }
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = forms[line].DocumentVid_DB.ToString(),
                Message = "При коде вида документа равном 19, в примечании к ячейке формы должно быть приведено наименование документа."
            });
        }
        return result;
    }

    #endregion

    #region Check041
    
    //Номер документа не пустой (колонка 17)
    private static List<CheckError> Check_041(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var valid = !string.IsNullOrWhiteSpace(forms[line].DocumentNumber_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "DocumentNumber_DB",
                Value = forms[line].DocumentNumber_DB,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check042
    
    //Дата документа <= дате операции для всех кодов операции, кроме 10 и 41
    private static List<CheckError> Check_042(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] excludedOperationCodes = { "10", "41" };
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        var documentDate = forms[line].DocumentDate_DB;
        if (excludedOperationCodes.Contains(operationCode)) return result;
        var valid = DateTime.TryParse(documentDate, out var documentDateReal)
                    && DateTime.TryParse(operationDate, out var operationDateReal)
                    && documentDateReal <= operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата документа не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check043
    
    //Дата документа = дате операции, если код операции 41
    private static List<CheckError> Check_043(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "41" };
        var documentDate = forms[line].DocumentDate_DB;
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = DateTime.TryParse(documentDate, out var documentDateReal)
                    && DateTime.TryParse(operationDate, out var operationDateReal)
                    && documentDateReal == operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата документа должна соответствовать дате операции."
            });
        }
        return result;
    }

    #endregion

    #region Check044
    
    //При коде операции 10, отчетный период должен оканчиваться не позднее 10 дней от дата документа (колонка 18)
    private static List<CheckError> Check_044(List<Form11> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10" };
        var documentDate = forms[line].DocumentDate_DB;
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = DateTime.TryParse(documentDate, out var documentDateReal)
                    && DateTime.TryParse(rep.StartPeriod_DB, out var dateBeginReal)
                    && DateTime.TryParse(rep.EndPeriod_DB, out var dateEndReal)
                    && documentDateReal >= dateBeginReal
                    && (dateEndReal - documentDateReal).Days is >= 0 and <= 10;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата документа выходит за границы периода."
            });
        }
        return result;
    }

    #endregion

    #region Check045
    
    //Код ОКПО поставщика/получателя 8/14 чисел и равен ОКПО отчитывающейся организации (колонка 19)
    private static List<CheckError> Check_045(List<Form11> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10", "11", "12", "15", "17", "18", "41", "42", "43", "46", "53", "58", "61", "62", "65", "67", "68", "71", "72", "73", "74", "75", "76" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB) 
            ? forms10[1].Okpo_DB 
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO)
                     && providerOrRecieverOKPO == repOKPO;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check046
    
    //Код ОКПО поставщика/получателя состоит из 8/14 чисел для определенных кодов операции (колонка 19)
    private static List<CheckError> Check_046(List<Form11> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "21", "25", "27", "28", "29", "31", "35", "37", "38", "39", "54", "63", "64", "66" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var okpoRep = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB) 
            ? forms10[1].Okpo_DB 
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        if (okpoRep == providerOrRecieverOKPO)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOKPO,
                Message = "Код ОКПО поставщика или получателя не должен совпадать с ОКПО отчитывающейся организации"
            });
        }
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check047
    
    //Код ОКПО поставщика/получателя состоит из 8/14 чисел или "минобороны" для определенных кодов операции (колонка 19)
    private static List<CheckError> Check_047(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "22", "32" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO) 
                    || providerOrRecieverOKPO.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check048
    
    //Код ОКПО поставщика/получателя из ОКСМ (не Россия), для определенных кодов операции, с примечанием (колонка 19)
    private static List<CheckError> Check_048(List<Form11> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "81", "82", "83", "84", "85", "86", "87", "88" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OKSM.Any(oksmEntry => oksmEntry["shortname"] == providerOrRecieverOKPO) 
                    && !providerOrRecieverOKPO.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        valid = false;
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
                    && note.GraphNumber_DB == 19.ToString() 
                    && !string.IsNullOrWhiteSpace(note.Comment_DB)))
            {
                valid = true;
                break;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOKPO,
                Message = "Необходимо добавить примечание."
            });
        }
        return result;
    }

    #endregion

    #region Check049
    
    //Код ОКПО поставщика/получателя состоит из 8/14 чисел для определенных кодов операции (колонка 19)
    private static List<CheckError> Check_049(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "97", "98", "99" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check050
    
    //При определенных кодах операции, код ОКПО перевозчика равен "-" (колонка 20)
    private static List<CheckError> Check_050(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10", "11", "12", "15", "17", "18", "41", "42", "43", "46", "53", "54", "58", "65", "66", "67", "68", "71", "72", "73", "74", "75", "76", "97", "98" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        var valid = transporterOKPO is "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = transporterOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check051

    //При определенных кодах операции, код ОКПО перевозчика равен 8/14 цифр (колонка 20)
    private static List<CheckError> Check_051(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "21", "25", "27", "28", "29", "31", "32", "35", "36", "37", "38", "39", "61", "62", "81", "82", "83", "84", "85", "86", "87", "88" };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(transporterOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = transporterOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check052
    
    //Код ОКПО перевозчика состоит из 8/14 чисел или "минобороны" для определенных кодов операции (колонка 20)
    private static List<CheckError> Check_052(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "22", "32" };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = okpoRegex.IsMatch(transporterOKPO) 
                    || transporterOKPO.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = transporterOKPO,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check053
    
    //Не пустое поле (колонка 21)
    private static List<CheckError> Check_053(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var packName = forms[line].PackName_DB;
        var valid = !string.IsNullOrWhiteSpace(packName);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "PackName_DB",
                Value = packName,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check054
    
    //Не пустое поле (колонка 22)
    private static List<CheckError> Check_054(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var packType = forms[line].PackType_DB;
        var valid = !string.IsNullOrWhiteSpace(packType);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "PackType_DB",
                Value = packType,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #region Check055
    
    //Не пустое поле (колонка 23)
    private static List<CheckError> Check_055(List<Form11> forms, int line)
    {
        List<CheckError> result = new();
        var packNumber = forms[line].PackNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(packNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_11",
                Row = (line + 1).ToString(),
                Column = "PackNumber_DB",
                Value = packNumber,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        return result;
    }

    #endregion

    #endregion
}