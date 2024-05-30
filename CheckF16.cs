using Models.CheckForm;
using Models.Collections;
using Models.Forms;
using Models.Forms.Form1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Client_App.Commands.SyncCommands.CheckForm;

public class CheckF16 : CheckBase
{
    #region Properties

    private static readonly string[] OperationCode_DB_Valids =
    {
        "10","11","12","13","14","16","18",
        "21","22","25","26","27","28","29",
        "31","32","35","36","37","38","39",
        "42","43","44","45","48","49","51",
        "52","56","57","59","63","64","68"
    };

    private static readonly Dictionary<string, string> GraphsList = new()
    {
        { "NumberInOrder_DB", "01 - № п/п" },
        { "OperationCode_DB", "02 - Код операции" },
        { "OperationDate_DB", "03 - Дата операции" },
        { "CodeRAO_DB", "04 - Код РАО" },
        { "StatusRAO_DB", "05 - Статус РАО" },
        { "Volume_DB", "06 - Объем без упаковки, куб. м" },
        { "Mass_DB", "07 - масса без упаковки (нетто), т" },
        { "QuantityOZIII_DB", "08 - Количество ОЗИИИ, шт." },
        { "MainRadionuclids_DB", "09 - Основные радионуклиды" },
        { "TritiumActivity_DB", "10 - Суммарная активность, Бк: тритий" },
        { "BetaGammaActivity_DB", "11 - Суммарная активность, Бк: бета-, гамма-излучающие радионуклиды" },
        { "AlphaActivity_DB", "12 - Суммарная активность, Бк: альфа-излучающие радионуклиды" },
        { "TransuraniumActivity_DB", "13 - Суммарная активность, Бк: трансурановые радионуклиды" },
        { "ActivityMeasurementDate_DB", "14 - Дата измерения активности" },
        { "DocumentVid_DB", "15 - Вид документа" },
        { "DocumentNumber_DB", "16 - Номер документа" },
        { "DocumentDate_DB", "17 - Дата документа" },
        { "ProviderOrRecieverOKPO_DB", "18 - Код ОКПО поставщика или получателя" },
        { "TransporterOKPO_DB", "19 - Код ОКПО перевозчика" },
        { "StoragePlaceName_DB", "20 - Наименование пункта хранения" },
        { "StoragePlaceCode_DB", "21 - Код пункта хранения" },
        { "RefineOrSortRAOCode_DB", "22 - Код переработки/сортировки РАО" },
        { "PackName_DB", "23 - Наименование УКТ, упаковки или иной ученой единицы" },
        { "PackType_DB", "24 - Тип УКТ, упаковки или иной ученой единицы" },
        { "PackNumber_DB", "25 - Заводской номер УКТ, упаковки или иной ученой единицы" },
        { "Subsidy_DB", "26 - Субсидия, %" },
        { "FcpNumber_DB", "27 - Номер мероприятия ФЦП" },
    };

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
        if (R.Count == 0)
        {
#if DEBUG
            R_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\..\..\")), "data", "Spravochniki", "R.xlsx"));
#else
            R_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"R.xlsx"));
#endif
        }
        if (HolidaysSpecific.Count == 0)
        {
#if DEBUG
            Holidays_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\..\\")), "data", "Spravochniki", "Holidays.xlsx"));
#else
            Holidays_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"Holidays.xlsx"));
#endif
        }
        foreach (var key in rep.Rows16)
        {
            var form = (Form16)key;
            var formsList = rep.Rows16.ToList<Form16>();
            var notes = rep.Notes.ToList<Note>();
            var forms10 = reps.Master_DB.Rows10.ToList<Form10>();
            errorList.AddRange(Check_001(formsList, currentFormLine));
            errorList.AddRange(Check_002(formsList, currentFormLine));
            errorList.AddRange(Check_002_11(formsList, currentFormLine));
            errorList.AddRange(Check_002_12(formsList, currentFormLine));
            errorList.AddRange(Check_002_29(formsList, notes, currentFormLine));
            errorList.AddRange(Check_002_51(formsList, currentFormLine));
            errorList.AddRange(Check_002_52(formsList, currentFormLine));
            errorList.AddRange(Check_002_57(formsList, currentFormLine));
            errorList.AddRange(Check_002_10(formsList, currentFormLine));
            errorList.AddRange(Check_002_49(formsList, currentFormLine));
            errorList.AddRange(Check_002_71(formsList, currentFormLine));
            errorList.AddRange(Check_003_non10(formsList, rep, currentFormLine));
            errorList.AddRange(Check_003_10(formsList, rep, currentFormLine));
            errorList.AddRange(Check_004(formsList, notes, currentFormLine));
            errorList.AddRange(Check_005_11(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_005_76(formsList, currentFormLine));
            errorList.AddRange(Check_005_Other(formsList, currentFormLine));
            errorList.AddRange(Check_006(formsList, currentFormLine));
            errorList.AddRange(Check_007(formsList, currentFormLine));
            errorList.AddRange(Check_008(formsList, currentFormLine));
            errorList.AddRange(Check_009(formsList, currentFormLine));
            errorList.AddRange(Check_010(formsList, currentFormLine));
            errorList.AddRange(Check_011(formsList, currentFormLine));
            errorList.AddRange(Check_012(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_013(formsList, currentFormLine));
            errorList.AddRange(Check_014(formsList, currentFormLine));
            errorList.AddRange(Check_015(formsList, currentFormLine));
            errorList.AddRange(Check_016(formsList, notes, currentFormLine));
            errorList.AddRange(Check_017(formsList, rep, currentFormLine));
            errorList.AddRange(Check_018_10(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_018_21(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_018_22(formsList, currentFormLine));
            errorList.AddRange(Check_019_01(formsList, currentFormLine));
            errorList.AddRange(Check_019_21(formsList, currentFormLine));
            errorList.AddRange(Check_019_22(formsList, currentFormLine));
            errorList.AddRange(Check_020(formsList, currentFormLine));
            errorList.AddRange(Check_021(formsList, currentFormLine));
            errorList.AddRange(Check_022_44(formsList, currentFormLine));
            errorList.AddRange(Check_022_49(formsList, currentFormLine));
            errorList.AddRange(Check_022_10(formsList, currentFormLine));
            errorList.AddRange(Check_022_22_44(formsList, currentFormLine));
            errorList.AddRange(Check_022_22_56(formsList, currentFormLine));
            errorList.AddRange(Check_023(formsList, currentFormLine));
            errorList.AddRange(Check_024(formsList, currentFormLine));
            errorList.AddRange(Check_025(formsList, currentFormLine));
            errorList.AddRange(Check_026(formsList, currentFormLine));
            errorList.AddRange(Check_027(formsList, currentFormLine));
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

    private static List<CheckError> Check_001(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].NumberInOrder_DB != line + 1)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "NumberInOrder_DB",
                Value = forms[line].NumberInOrder_DB.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Номера строк должны располагаться по порядку, без пропусков или дублирования номеров"
            });
        }
        return result;
    }

    #endregion

    #region Check002

    private static List<CheckError> Check_002(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var valid = operationCode != null && OperationCode_DB_Valids.Contains(operationCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Код операции не может быть использован в форме 1.6."
            });
        }
        return result;
    }

    #endregion

    #region Check002_11

    private static List<CheckError> Check_002_11(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "11", "12", "13", "14", "16", "18", "41", "56", "57", "59" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        if (forms[line].CodeRAO_DB.Length < 10) return result;
        var valid = forms[line].CodeRAO_DB.Substring(8, 2) != "99";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для вновь образованных РАО код типа РАО «99» не может быть использован"
            });
        }
        return result;
    }

    #endregion

    #region Check002_12

    private static List<CheckError> Check_002_12(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "12,42" };
        var requiredNuclids = new string[] { "плутоний", "уран-233", "уран-235", "уран-238", "нептуний-237", "америций-241", "америций-243" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var nuclids = forms[line].MainRadionuclids_DB.Split(';');
        var valid = false;
        for (var i = 0; i < nuclids.Length; i++)
        {
            nuclids[i] = nuclids[i].Trim().ToLower();
            if (requiredNuclids.Contains(nuclids[i]))
            {
                valid = true;
                break;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "MainRadionuclids_DB",
                Value = forms[line].MainRadionuclids_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В графе 9 не представлены сведения о радионуклидах, которые могут быть отнесены к ЯМ. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_29

    private static List<CheckError> Check_002_29(List<Form16> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "29", "39", "49", "59", "97", "98", "99" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        const byte graphNumber = 2;
        var valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо дать пояснение об осуществленной операции."
            });
        }
        return result;
    }

    #endregion

    #region Check002_51

    private static List<CheckError> Check_002_51(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "51" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + ""
            });
        }
        return result;
    }

    #endregion

    #region Check002_52

    private static List<CheckError> Check_002_52(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "52" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В отчетах не найдена строка об осуществлении передачи учетной единицы. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_57
    private static List<CheckError> Check_002_57(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "57" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В отчетах не найдена строка снятии учетной единицы для упаковки/переупаковки. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_10
    private static List<CheckError> Check_002_10(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "10", "11", "12", "13", "14", "16", "18", "31", "32", "35", "36", "37", "38", "41", "56", "57", "59", "97" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "По сведениям, представленным в строке, отходы не относятся к РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check002_49
    private static List<CheckError> Check_002_49(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "49", "59" };
        var graph22ValidValues = new string[] { "-", "52", "72", "74" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = graph22ValidValues.Contains(forms[line].RefineOrSortRAOCode_DB.Trim());
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Код операции не соответствует коду переработки/сортировки, указанному в графе 22"
            });
        }
        return result;
    }

    #endregion

    #region Check002_71
    //Справочная "ошибка" - т.е. даже не ошибка.
    private static List<CheckError> Check_002_71(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "71", "72", "73", "74", "75", "76", "18", "68" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "К отчету необходимо приложить скан-копию документа характеризующего операцию"
            });
        }
        return result;
    }

    #endregion

    #region Check003_non10
    private static List<CheckError> Check_003_non10(List<Form16> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] nonApplicableOperationCodes = { "10" };
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        if (nonApplicableOperationCodes.Contains(operationCode)) return result;
        var valid = operationDate != null;
        DateTime pEnd;
        DateTime pMid;
        if (valid && rep is { StartPeriod_DB: not null, EndPeriod_DB: not null })
        {
            valid = DateTime.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateTime.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateTime.TryParse(operationDate, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(operationDate),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата операции не входит в отчетный период."
            });
        }
        return result;
    }

    #endregion

    #region Check003_10
    private static List<CheckError> Check_003_10(List<Form16> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10" };
        var operationCode = forms[line].OperationCode_DB;
        var documentDate = forms[line].DocumentDate_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = documentDate != null;
        DateTime pEnd;
        DateTime pMid;
        if (valid && rep is { StartPeriod_DB: not null, EndPeriod_DB: not null })
        {
            valid = DateTime.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateTime.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateTime.TryParse(documentDate, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(documentDate),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата документа не входит в отчетный период."
            });
        }
        return result;
    }

    #endregion

    #region Check004
    private static List<CheckError> Check_004(List<Form16> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var CodeRAO_DB = forms[line].CodeRAO_DB;
        var valid = !string.IsNullOrWhiteSpace(CodeRAO_DB) && CodeRAO_DB.Length == 11;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность заполнения кода РАО."
            });
        }
        return result;
    }

    #endregion

    #region Check005_11
    private static List<CheckError> Check_005_11(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "11", "12", "13", "14", "41" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var repOKPO_List = forms10.Select(x => x.Okpo_DB).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        var valid = repOKPO_List.Contains(StatusRAO_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При данном коде операции, статус РАО должен быть равен коду ОКПО отчитывающейся организации или юридического лица (РАО, образовавшиеся после 15.07.2011, находятся в собственности организации, в результате деятельности которой они образовались)."
            });
        }
        return result;
    }

    #endregion

    #region Check005_26
    private static List<CheckError> Check_005_26(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "26", "28", "63" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var repOKPO_List = forms10.Select(x => x.Okpo_DB).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        var valid = repOKPO_List.Contains(StatusRAO_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При данном коде операции, статус РАО должен быть равен коду ОКПО отчитывающейся организации или юридического лица (операция, соответствующая выбранному коду, может использоватся только для собственных РАО)."
            });
        }
        return result;
    }

    #endregion

    #region Check005_38
    private static List<CheckError> Check_005_38(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "38", "64" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var repOKPO_List = forms10.Select(x => x.Okpo_DB).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        var valid = repOKPO_List.Contains(StatusRAO_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При операциях, связанных с получением права собственности, в графе статус РАО необходимо отразить код ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check005_42
    private static List<CheckError> Check_005_42(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "42", "43", "73", "97", "98" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var repOKPO_List = forms10.Select(x => x.Okpo_DB).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        var valid = repOKPO_List.Contains(StatusRAO_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При данном коде операции, статус РАО должен быть равен коду ОКПО отчитывающейся организации или юридического лица"
            });
        }
        return result;
    }

    #endregion

    #region Check005_22
    private static List<CheckError> Check_005_22(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "22", "23" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(StatusRAO_DB) || StatusRAO_DB == "1";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность статуса РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check005_16
    private static List<CheckError> Check_005_16(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "16" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(StatusRAO_DB) || StatusRAO_DB == "2";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность статуса РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check005_76
    private static List<CheckError> Check_005_76(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "76" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var valid = StatusRAO_DB is "6" or "9";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность статуса РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check005_Other

    private static List<CheckError> Check_005_Other(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "10", "18", "21", "25", "27", "29", "31", "35", "36", "37", "39", "44", "45", "48", "49", "51", "52", "56", "57", "59", "68", "71", "74", "75" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var applicableRAOStatuses = new string[] { "1", "2", "3", "4", "6", "9" };
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(StatusRAO_DB) || applicableRAOStatuses.Contains(StatusRAO_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполнение графы 5 не соответствует требованиям приказа Госкорпорации \"Росатом\" от 07.12.2020 № 1/13-НПА"
            });
        }
        return result;
    }

    #endregion

    #region Check006
    private static List<CheckError> Check_006(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var Volume_DB = forms[line].Volume_DB;
        var valid = float.TryParse(ConvertStringToExponential(Volume_DB),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out var value) && value > 0;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "Volume_DB",
                Value = Volume_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить сведения об объеме РАО, если представляемые данные являются расчетными, то соответствующие значения указываются в круглых скобках"
            });
        }
        return result;
    }

    #endregion

    #region Check007
    private static List<CheckError> Check_007(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var Mass_DB = forms[line].Mass_DB;
        var Volume_DB = forms[line].Volume_DB;
        float Mass_Real;
        float Volume_Real;
        var valid = float.TryParse(ConvertStringToExponential(Mass_DB),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out Mass_Real) && Mass_Real > 0;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = Mass_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить сведения о массе РАО, если представляемые данные являются расчетными, то соответствующие значения указываются в круглых скобках"
            });
            return result;
        }
        if (!float.TryParse(ConvertStringToExponential(Volume_DB), out Volume_Real)) return result;
        if (Volume_Real == 0.0) return result;
        var Density_Real = Mass_Real / Volume_Real;
        if (Density_Real > 21.0)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = Density_Real.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение массы и объема. Расчетное значение плотности слишком большое"
            });
        }
        else if (Density_Real < 0.005)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = Density_Real.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение массы и объема. Расчетное значение плотности слишком маленькое"
            });
        }
        return result;
    }

    #endregion

    #region Check008
    private static List<CheckError> Check_008(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var CodeRAO_DB = forms[line].CodeRAO_DB.Trim();
        var QuantityOZIII_DB = forms[line].QuantityOZIII_DB.Trim();
        int QuantityOZIII_Real = -1;
        int.TryParse(QuantityOZIII_DB, out QuantityOZIII_Real);
        var RAOTypes_1 = new string[] { "81", "82", "84", "85", "86", "87", "88", "89" };
        var RAOTypes_2 = new string[] { "99" };
        if (CodeRAO_DB.Length < 10) return result;
        var TypeRAO_DB = CodeRAO_DB.Substring(8, 2);
        var StateRAO_DB = CodeRAO_DB.Substring(0, 1);
        if (RAOTypes_1.Contains(TypeRAO_DB))
        {
            if (QuantityOZIII_Real == -1)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "QuantityOZIII_DB",
                    Value = QuantityOZIII_DB,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить сведения о количестве ОЗИИИ"
                });
            }
        }
        else if (RAOTypes_2.Contains(TypeRAO_DB))
        {
            if (QuantityOZIII_Real == -1 && QuantityOZIII_DB != "-")
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "QuantityOZIII_DB",
                    Value = QuantityOZIII_DB,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить сведения о количестве ОЗИИИ"
                });
            }
        }
        else
        {
            if (QuantityOZIII_DB != "-")
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "QuantityOZIII_DB",
                    Value = QuantityOZIII_DB,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа заполняется только для ОЗИИИ. Поставьте прочерк."
                });
            }
        }
        return result;
    }

    #endregion

    #region Check009
    private static List<CheckError> Check_009(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].MainRadionuclids_DB;
        if (string.IsNullOrWhiteSpace(radionuclids))
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните графу 9 «Основные радионуклиды»"
            });
            return result;
        }
        var CodeRAO_DB = forms[line].CodeRAO_DB.Trim();
        if (CodeRAO_DB.Length < 3) return result;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (radArray.Length == 1 && string.Equals(radArray[0], "-"))
        {
            if (CodeRAO_DB.Substring(2, 1) != "0")
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "Radionuclids_DB",
                    Value = Convert.ToString(radionuclids),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните графу 9 «Основные радионуклиды»"
                });
                return result;
            }
        }
        else
        {
            var valid = radArray.All(rad => R.Any(phEntry => phEntry["name"] == rad));
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "Radionuclids_DB",
                    Value = Convert.ToString(radionuclids),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните графу 9 «Основные радионуклиды»."
                });
            }
        }
        return result;
    }

    #endregion

    #region Check010
    private static List<CheckError> Check_010(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].TritiumActivity_DB;
        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (!radArray.Any(rad => R.Any(phEntry => phEntry["name"] == rad && phEntry["code"] == "т"))) return result;
        var activityReal = 1.0;
        if (!double.TryParse(ConvertStringToExponential(activity),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out activityReal))
        {
            if (activityReal is <= 10e+01 or > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "TritiumActivity_DB",
                    Value = Convert.ToString(forms[line].TritiumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности"
                });
            }
            else
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "TritiumActivity_DB",
                    Value = Convert.ToString(forms[line].TritiumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для указанного в графе 9 радионуклидного состава должна быть приведена активность в графе 10"
                });
            }
        }
        return result;
    }
    #endregion

    #region Check011
    private static List<CheckError> Check_011(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].BetaGammaActivity_DB;
        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (!radArray.Any(rad => R.Any(phEntry => phEntry["name"] == rad && phEntry["code"] == "б"))) return result;
        var activityReal = 1.0;
        if (!double.TryParse(ConvertStringToExponential(activity),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out activityReal))
        {
            if (activityReal is <= 10e+01 or > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "BetaGammaActivity_DB",
                    Value = Convert.ToString(forms[line].BetaGammaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности"
                });
            }
            else
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "BetaGammaActivity_DB",
                    Value = Convert.ToString(forms[line].BetaGammaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для указанного в графе 9 радионуклидного состава должна быть приведена активность в графе 11"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check012
    private static List<CheckError> Check_012(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].AlphaActivity_DB;
        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (!radArray.Any(rad => R.Any(phEntry => phEntry["name"] == rad && phEntry["code"] == "а"))) return result;
        var activityReal = 1.0;
        if (!double.TryParse(ConvertStringToExponential(activity),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out activityReal))
        {
            if (activityReal is <= 10e+01 or > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "AlphaActivity_DB",
                    Value = Convert.ToString(forms[line].AlphaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности"
                });
            }
            else
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "AlphaActivity_DB",
                    Value = Convert.ToString(forms[line].AlphaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для указанного в графе 9 радионуклидного состава должна быть приведена активность в графе 12"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check013
    private static List<CheckError> Check_013(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].TransuraniumActivity_DB;
        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (!radArray.Any(rad => R.Any(phEntry => phEntry["name"] == rad && phEntry["code"] == "у")))
            return result;
        if (!double.TryParse(ConvertStringToExponential(activity),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out var activityReal))
        {
            if (activityReal is <= 10e+01 or > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "TransuraniumActivity_DB",
                    Value = Convert.ToString(forms[line].TransuraniumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности"
                });
            }
            else
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "TransuraniumActivity_DB",
                    Value = Convert.ToString(forms[line].TransuraniumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для указанного в графе 9 радионуклидного состава должна быть приведена активность в графе 13"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check014
    private static List<CheckError> Check_014(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var activityDate = forms[line].ActivityMeasurementDate_DB;
        var operationDate = forms[line].OperationDate_DB;
        DateTime activityDateReal = DateTime.Now;
        DateTime operationDateReal = DateTime.UnixEpoch;
        var valid = DateTime.TryParse(activityDate, out activityDateReal) && DateTime.TryParse(operationDate, out operationDateReal);
        if (valid)
        {
            valid = activityDateReal <= operationDateReal;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "ActivityMeasurementDate_DB",
                Value = activityDate,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность указанной даты измерения активности"
            });
        }
        return result;
    }

    #endregion

    #region Check015
    private static List<CheckError> Check_015(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        byte?[] validDocument = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 19 };
        var DocumentVid_DB = forms[line].DocumentVid_DB;
        var valid = validDocument.Contains(DocumentVid_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = DocumentVid_DB.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность заполнения графы 15"
            });
        }
        return result;
    }

    #endregion

    #region Check016
    private static List<CheckError> Check_016(List<Form16> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var documentNumberDB = forms[line].DocumentNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(documentNumberDB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "DocumentNumber_DB",
                Value = Convert.ToString(documentNumberDB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check017
    private static List<CheckError> Check_017(List<Form16> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        var documentDate = forms[line].DocumentDate_DB;
        DateTime pMid;
        DateTime pEnd;
        var valid = true;
        if (operationCode == "41")
        {
            valid = DateTime.TryParse(documentDate, out pMid)
                    && DateTime.TryParse(operationDate, out var pOper)
                    && pMid.Date == pOper.Date;
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "DocumentDate_DB",
                    Value = Convert.ToString(documentDate),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата документа должна соответствовать дате операции"
                });
            }
        }
        else if (operationCode == "10")
        {
            valid = DateTime.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateTime.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateTime.TryParse(documentDate, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "DocumentDate_DB",
                    Value = Convert.ToString(documentDate),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата документа выходит за границы периода"
                });
            }
        }
        else
        {
            valid = DateTime.TryParse(documentDate, out pMid)
                    && DateTime.TryParse(operationDate, out var pOper)
                    && pMid.Date <= pOper.Date;
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "DocumentDate_DB",
                    Value = Convert.ToString(documentDate),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Даат документа не может быть позже даты операции"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check018_10

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_018_10(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10", "11", "12", "13", "14", "16", "18", "41", "42", "43", "44", "45", "48", "49", "51", "52", "56", "57", "59", "68", "71", "72", "73", "74", "75", "76", "97", "98" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = providerOrRecieverOKPO == repOKPO;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для выбранного кода операции указывается код ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check018_21

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_018_21(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "21", "25", "26", "27", "28", "29", "31", "35", "36", "37", "38", "39", "99" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Значение может состоять только из 8 или 14 символов"
            });
        }
        valid = providerOrRecieverOKPO != repOKPO;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для выбранного кода операции указывается код ОКПО контрагента."
            });
        }
        return result;
    }

    #endregion

    #region Check018_22

    //Код ОКПО поставщика/получателя состоит из 8/14 чисел или "минобороны" (колонка 18)
    private static List<CheckError> Check_018_22(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (operationCode is not ("22" or "32")) return result;
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO)
                    || providerOrRecieverOKPO.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Формат ввода данных не соответствует приказу. Следует указать код ОКПО контрагента, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check019_01

    //При определенных кодах операции, код ОКПО перевозчика равен "-"
    private static List<CheckError> Check_019_01(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = {
            "01", "10", "11","12","13","14","16", "18", "41", "42","43", "44", "45", "48", "49", "51", "52","56", "57", "59", "71", "72", "73", "74", "75",
            "76", "97", "98", "99"
        };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        var valid = transporterOKPO is "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При выбранном коде операции транспортирование не производится."
            });
        }
        return result;
    }

    #endregion

    #region Check019_21

    //При определенных кодах операции, код ОКПО перевозчика равен 8/14 цифр
    private static List<CheckError> Check_019_21(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21","25","26","27","28","29","31","35","36","37","38","39"
        };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (okpoRegex.IsMatch(transporterOKPO)) return result;

        var valid = okpoRegex.IsMatch(transporterOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Значение может состоять только из 8 или 14 символов"
            });
        }
        return result;
    }

    #endregion

    #region Check019_22

    //Код ОКПО перевозчика состоит из 8/14 чисел или "минобороны"
    private static List<CheckError> Check_019_22(List<Form16> forms, int line)
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
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо указать код ОКПО организации перевозчика, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check020
    private static List<CheckError> Check_020(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var StoragePlaceName_DB = forms[line].StoragePlaceName_DB;
        var valid = !string.IsNullOrWhiteSpace(StoragePlaceName_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StoragePlaceName_DB",
                Value = Convert.ToString(StoragePlaceName_DB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check021
    private static List<CheckError> Check_021(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var StoragePlaceCode_DB = forms[line].StoragePlaceCode_DB;
        var valid = !string.IsNullOrWhiteSpace(StoragePlaceCode_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "StoragePlaceCode_DB",
                Value = Convert.ToString(StoragePlaceCode_DB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check022_44
    private static List<CheckError> Check_022_44(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "44", "45", "49", "56", "57", "59" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check022_49
    private static List<CheckError> Check_022_49(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "49", "59" };
        var applicableRefineCodes = new string[] { "52", "72", "74" };
        var operationCode = forms[line].OperationCode_DB;
        var RefineOrSortRAOCode_DB = forms[line].RefineOrSortRAOCode_DB;
        if (!applicableRefineCodes.Contains(RefineOrSortRAOCode_DB)) return result;
        var valid = applicableOperationCodes.Contains(operationCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для выбранных кодов способа сортировки, необходимо выбрать код операции 49 или 59"
            });
        }
        return result;
    }

    #endregion

    #region Check022_10
    private static List<CheckError> Check_022_10(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "10","11","12","13","14","16","18","21","22","25","26","27","28","29","31","32","35",
            "36","37","38","39","42","43","48","51","52","63","64","68","71","72","73","74","75",
            "76","97","98","99" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = forms[line].RefineOrSortRAOCode_DB.Trim() == "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При данном коде операции для кода переработки/сортировки используется символ \"-\""
            });
        }
        return result;
    }

    #endregion

    #region Check022_22_44
    private static List<CheckError> Check_022_22_44(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].RefineOrSortRAOCode_DB != "22") return result;
        var CodeRAO_DB = forms[line].CodeRAO_DB;
        if (CodeRAO_DB.Length < 11) return result;
        if (forms[line].OperationCode_DB != "44") return result;
        var valid = CodeRAO_DB.Substring(10, 1) == "1";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение 11 символа кода РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check022_22_56
    private static List<CheckError> Check_022_22_56(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].RefineOrSortRAOCode_DB != "22") return result;
        var CodeRAO_DB = forms[line].CodeRAO_DB;
        if (CodeRAO_DB.Length < 10) return result;
        if (forms[line].OperationCode_DB != "56") return result;
        var valid = (CodeRAO_DB.Substring(0, 1) is "2" && CodeRAO_DB.Substring(8, 2) is "66" or "74") || (CodeRAO_DB.Substring(0, 1) is "1" && CodeRAO_DB.Substring(8, 2) is "14");
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "РАО направлены на переработку на установке сжигания. Проверьте значение 11 символа кода РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check023
    private static List<CheckError> Check_023(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackName_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "PackName_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check024
    private static List<CheckError> Check_024(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackType_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "PackType_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check025
    private static List<CheckError> Check_025(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "PackNumber_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check026
    private static List<CheckError> Check_026(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].Subsidy_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "Subsidy_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
            return result;
        }
        if (float.TryParse(field_value, out var value_real))
        {
            if (value_real < 0 || value_real > 100)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = (line + 1).ToString(),
                    Column = "Subsidy_DB",
                    Value = Convert.ToString(field_value),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение субсидии."
                });
            }
        }
        return result;
    }

    #endregion

    #region Check027
    private static List<CheckError> Check_027(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].FcpNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = (line + 1).ToString(),
                Column = "FcpNumber_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #endregion
}
