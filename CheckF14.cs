using Models.CheckForm;
using Models.Collections;
using Models.Forms.Form1;
using System;
using System.Collections.Generic;
using System.IO;
using Models.Forms;
using System.Linq;
using System.Globalization;
using FirebirdSql.EntityFrameworkCore.Firebird.Query.ExpressionTranslators.Internal;
using System.Text.RegularExpressions;
using System.Reflection;

namespace Client_App.Commands.SyncCommands.CheckForm;

public abstract class CheckF14 : CheckBase
{
    #region Check_Total

    public static List<CheckError> Check_Total(Reports reps, Report rep)
    {
        var currentFormLine = 0;
        List<CheckError> errorList = new();

        if (OKSM.Count == 0)
        {
#if DEBUG
            OKSM_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\..\..\")), "data", "Spravochniki", "oksm.xlsx"));
#else
            OKSM_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"oksm.xlsx"));
#endif
        }
        if (D.Count == 0)
        {
#if DEBUG
            D_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\..\..\")), "data", "Spravochniki", "D.xlsx"));
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
        var formsList = rep.Rows14.ToList<Form14>();
        foreach (var key in rep.Rows14)
        {
            var form = (Form14)key;
            var notes = rep.Notes.ToList<Note>();
            var forms10 = reps.Master_DB.Rows10.ToList<Form10>();
            errorList.AddRange(Check_001(formsList, currentFormLine));
            errorList.AddRange(Check_002(formsList, currentFormLine));
            errorList.AddRange(Check_003(formsList, currentFormLine));
            errorList.AddRange(Check_004(formsList, currentFormLine));
            errorList.AddRange(Check_005(formsList, notes, currentFormLine));
            errorList.AddRange(Check_006(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_007(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_008(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_009(formsList, currentFormLine));
            errorList.AddRange(Check_010(formsList, currentFormLine));
            errorList.AddRange(Check_011(formsList, currentFormLine));
            errorList.AddRange(Check_012(formsList, currentFormLine));
            errorList.AddRange(Check_013(formsList, currentFormLine));
            errorList.AddRange(Check_014(formsList, currentFormLine));
            errorList.AddRange(Check_015(formsList, currentFormLine));
            errorList.AddRange(Check_016(formsList, currentFormLine));
            errorList.AddRange(Check_017(formsList, currentFormLine));
            errorList.AddRange(Check_018(formsList, rep, currentFormLine));
            errorList.AddRange(Check_019(formsList, rep, currentFormLine));
            errorList.AddRange(Check_020(formsList, currentFormLine));
            errorList.AddRange(Check_021(formsList, currentFormLine));
            errorList.AddRange(Check_022(formsList, currentFormLine));
            errorList.AddRange(Check_023(formsList, currentFormLine));
            errorList.AddRange(Check_024(formsList, currentFormLine));
            errorList.AddRange(Check_025(formsList, currentFormLine));
            errorList.AddRange(Check_026(formsList, currentFormLine));
            errorList.AddRange(Check_027(formsList, currentFormLine));
            errorList.AddRange(Check_028(formsList, currentFormLine));
            errorList.AddRange(Check_029(formsList, currentFormLine));
            errorList.AddRange(Check_030(formsList, currentFormLine));
            errorList.AddRange(Check_031(formsList, currentFormLine));
            errorList.AddRange(Check_032(formsList, currentFormLine));
            errorList.AddRange(Check_033(formsList, currentFormLine));
            errorList.AddRange(Check_034(formsList, currentFormLine));
            errorList.AddRange(Check_035(formsList, currentFormLine));
            errorList.AddRange(Check_036(formsList, currentFormLine));
            errorList.AddRange(Check_037(formsList, currentFormLine));
            errorList.AddRange(Check_038(formsList, notes, currentFormLine));
            errorList.AddRange(Check_039(formsList, notes, currentFormLine));
            errorList.AddRange(Check_040(formsList, notes, currentFormLine));
            errorList.AddRange(Check_041(formsList, currentFormLine));
            errorList.AddRange(Check_042(formsList, currentFormLine));
            errorList.AddRange(Check_043(formsList, notes, currentFormLine));
            errorList.AddRange(Check_044(formsList, currentFormLine));
            errorList.AddRange(Check_045(formsList, currentFormLine));
            errorList.AddRange(Check_046(formsList, currentFormLine));
            errorList.AddRange(Check_047(formsList, currentFormLine));
            errorList.AddRange(Check_048(formsList, rep, currentFormLine));
            errorList.AddRange(Check_049(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_050(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_051(formsList, currentFormLine));
            errorList.AddRange(Check_052(formsList, notes, currentFormLine));
            errorList.AddRange(Check_053(formsList, currentFormLine));
            errorList.AddRange(Check_054(formsList, notes, currentFormLine));
            errorList.AddRange(Check_055(formsList, currentFormLine));
            errorList.AddRange(Check_056(formsList, currentFormLine));
            errorList.AddRange(Check_057(formsList, currentFormLine));
            errorList.AddRange(Check_058(formsList, currentFormLine));
            currentFormLine++;
        }
        errorList.AddRange(Check_059(formsList));
        var index = 0;
        foreach (var error in errorList)
        {
            error.Column = GraphsList.GetValueOrDefault(error.Column, error.Column);
            index++;
            error.Index = index;
        }
        return errorList;
    }

    #endregion

    #region Checks

    #region Check001

    private static List<CheckError> Check_001(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        return result;
    }

    #endregion

    #region Check002

    private static List<CheckError> Check_002(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].NumberInOrder_DB != line + 1)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "NumberInOrder_DB",
                Value = forms[line].Id.ToString(),
                Message = "Номера строк должны располагаться по порядку, без пропусков или дублирования номеров"
            });
        }
        return result;
    }

    #endregion

    #region Check003

    //Код операции из списка (колонка 2)
    private static List<CheckError> Check_003(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var operationCodeValid = new[]
        {
            "10", "11", "12", "15", "17", "18", "21", "22", "25", "27", "28", "29", "31",
            "32", "35", "37", "38", "39", "41", "42", "43", "46", "47", "48", "53", "54",
            "58", "61", "62", "63", "64", "65", "67", "68", "71", "72", "73", "74", "75",
            "81", "82", "83", "84", "85", "86", "87", "88", "97", "98", "99"
        };
        var valid = operationCode != null && operationCodeValid.Contains(operationCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = Convert.ToString(operationCode),
                Message = $"Код операции {operationCode} не может быть использован в форме 1.4."
            });
        }
        return result;
    }

    #endregion

    #region Check004

    //Для кодов операции 12, 42, в радионуклидах должен быть указан хоть один из списка (колонка 6)
    private static List<CheckError> Check_004(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var radionuclid = forms[line].Radionuclids_DB;
        string[] applicableOperationCodes = { "12", "42" };
        var radionuclidValid = new[]
        {
            "плутоний-238", "плутоний-239", "плутоний-240", "уран-233", "уран-235", "уран-238", "нептуний-237",
            "америций-241", "америций 243", "калифорний-252", "торий-232", "литий-6", "дейтерий", "тритий"
        };
        if (string.IsNullOrWhiteSpace(radionuclid) || !applicableOperationCodes.Contains(operationCode)) return result;
        var valid = radionuclid != null && radionuclidValid.Any(nuclid => radionuclid.ToLower().Contains(nuclid));
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclid),
                Message = "В графе 6 не представлены сведения о радионуклидах, которые могут быть отнесены к ЯМ. Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check005

    //Для определенных кодов операции должно быть примечание (колонка 2)
    private static List<CheckError> Check_005(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 2;
        var operationCode = forms[line].OperationCode_DB;
        string[] applicableOperationCodes = { "29", "39", "97", "98", "99" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = Convert.ToString(operationCode),
                Message = "Необходимо дать пояснение об осуществленной операции."
            });
        }
        return result;
    }

    #endregion

    #region Check006

    //Код ОКПО поставщика/получателя равен коду ОКПО отчитывающейся организации (колонка 18)
    private static List<CheckError> Check_006(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "53" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = providerOrRecieverOKPO.Equals(repOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = "В графе 18 необходимо указать код ОКПО своей организации. В случае, если зарядку/разрядку осуществляла подрядная организация, следует использовать код операции 54."
            });
        }
        return result;
    }

    #endregion

    #region Check007

    private static List<CheckError> Check_007(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "11", "12", "15", "28", "38", "41", "48", "63", "64", "65", "73", "81", "85", "88" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var owner = forms[line].Owner_DB;
        var okpoRepJur = forms10[0].Okpo_DB ?? "";
        var okpoRepTerPodr = forms10[1].Okpo_DB ?? "";
        var valid = !string.IsNullOrWhiteSpace(owner) 
                    && (owner == okpoRepTerPodr || owner == okpoRepJur);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Уточните правообладателя ОРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check008

    //Код ОКПО поставщика/получателя равен коду ОКПО отчитывающейся организации (колонка 18)
    private static List<CheckError> Check_008(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "54" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = !string.IsNullOrWhiteSpace(providerOrRecieverOKPO) && !providerOrRecieverOKPO.Equals(repOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = "В графе 18 необходимо указать ОКПО подрядной организации."
            });
        }
        return result;
    }

    #endregion

    #region Check009

    private static List<CheckError> Check_009(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check010

    private static List<CheckError> Check_010(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check011

    private static List<CheckError> Check_011(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check012

    private static List<CheckError> Check_012(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check013

    private static List<CheckError> Check_013(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check014

    private static List<CheckError> Check_014(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check015

    private static List<CheckError> Check_015(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check016

    private static List<CheckError> Check_016(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check017

    private static List<CheckError> Check_017(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check018

    //Дата документа входит в отчетный период с учетом срока подачи отчета в днях (колонка 3)
    private static List<CheckError> Check_018(List<Form14> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] nonApplicableOperationCodes = { "10" };
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        if (nonApplicableOperationCodes.Contains(operationCode)) return result;
        var valid = operationDate != null;
        var pEnd = DateOnly.MinValue;
        var pMid = DateOnly.MinValue;
        if (valid && rep is { StartPeriod_DB: not null, EndPeriod_DB: not null })
        {
            valid = DateOnly.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateOnly.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateOnly.TryParse(operationDate, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(operationDate),
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
                "41", "42", "43", "46", "47", "48", "53", "54", "58", "61", "62", "63", "64", "65", "67", "68", "72",
                "81", "82", "83", "84", "85", "86", "87", "88", "97", "98", "99"
            };
            if (operationCodeWithDeadline10.Contains(operationCode) && WorkdaysBetweenDates(pMid, pEnd) > 10)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_14",
                    Row = (line + 1).ToString(),
                    Column = "OperationDate_DB",
                    Value = Convert.ToString(operationDate),
                    Message = "Дата окончания отчетного периода превышает дату операции более чем на 10 дней."
                });
            }
            else if (operationCodeWithDeadline5.Contains(operationCode) && WorkdaysBetweenDates(pMid, pEnd) > 5)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_14",
                    Row = (line + 1).ToString(),
                    Column = "OperationDate_DB",
                    Value = Convert.ToString(operationDate),
                    Message = "Дата окончания отчетного периода превышает дату операции более чем на 5 дней."
                });
            }
            else if (operationCodeWithDeadline1.Contains(operationCode) && WorkdaysBetweenDates(pMid, pEnd) > 1)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_14",
                    Row = (line + 1).ToString(),
                    Column = "OperationDate_DB",
                    Value = Convert.ToString(operationDate),
                    Message = "Дата окончания отчетного периода превышает дату операции более чем на 1 день."
                });
            }
        }
        return result;
    }

    #endregion

    #region Check019

    //При коде операции 10, дата документа должна попадать в отчетный период
    private static List<CheckError> Check_019(List<Form14> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10" };
        var documentDate = forms[line].DocumentDate_DB;
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = DateOnly.TryParse(documentDate, out var documentDateReal)
                    && DateOnly.TryParse(rep.StartPeriod_DB, out var dateBeginReal)
                    && DateOnly.TryParse(rep.EndPeriod_DB, out var dateEndReal)
                    && documentDateReal >= dateBeginReal && documentDateReal <= dateEndReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата акта инвентаризации выходит за границы отчетного периода."
            });
        }
        return result;
    }

    #endregion

    #region Check020

    //Номер паспорта не пустая строка (колонка 4)
    private static List<CheckError> Check_020(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var passportNumber = forms[line].PassportNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(passportNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PassportNumber_DB",
                Value = Convert.ToString(passportNumber),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check021

    //Наименование не пустая строка (колонка 5)
    private static List<CheckError> Check_021(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var name = forms[line].Name_DB;
        var valid = !string.IsNullOrWhiteSpace(name);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Name_DB",
                Value = Convert.ToString(name),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check022

    //Вид от 4 до 12 (колонка 6)
    private static List<CheckError> Check_022(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var sort = forms[line].Sort_DB;
        var valid = sort is >= 4 and <= 12;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Sort_DB",
                Value = Convert.ToString(sort),
                Message = "Формат ввода данных не соответствует приказу. Необходимо указать вид РВ в соответствии с таблицей 4 приложения №2 к приказу Госкорпорации \"Росатом\" от 07.12.2020 №1/13-НПА."
            });
        }
        return result;
    }

    #endregion

    #region Check023

    //Радионуклиды не пустая строка (колонка 6)
    private static List<CheckError> Check_023(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].Radionuclids_DB;
        var valid = !string.IsNullOrWhiteSpace(radionuclids);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check024

    //У радионуклидов в качестве разделителя использовать ; (колонка 6)
    private static List<CheckError> Check_024(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].Radionuclids_DB;
        if (radionuclids.Contains(',') || radionuclids.Contains('+'))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = "Формат ввода данных не соответствует приказу. Радионуклиды должны быть разделены точкой с запятой."
            });
        }
        return result;
    }

    #endregion

    #region Check025

    //Все радионуклиды есть в справочнике (колонка 6)
    private static List<CheckError> Check_025(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].Radionuclids_DB;
        if (string.IsNullOrWhiteSpace(radionuclids)) return result;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (radArray.Length == 1 && string.Equals(radArray[0], "-")) return result;
        var valid = radArray.All(rad => R.Any(phEntry => phEntry["name"] == rad));
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = "Формат ввода данных не соответствует приказу. Радионуклид отсутствует в справочнике."
            });
        }
        return result;
    }

    #endregion

    #region Check026

    //Каждый из радионуклидов не короткоживущий, >=60 суток (колонка 6)
    private static List<CheckError> Check_026(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].Radionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower().Split(';');
        var valid = true;
        var shortRad = "";
        double shortRadHalfLife = 0;
        foreach (var rad in radArray)
        {
            var phEntry = R.FirstOrDefault(phEntry => phEntry["name"] == rad);
            if (phEntry is null) return result;
            var unit = phEntry["unit"];
            var value = phEntry["value"].Replace('.', ',');
            if (!double.TryParse(value,
                    NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                    CultureInfo.CreateSpecificCulture("ru-RU"),
                    out var halfLife)) return result;
            switch (unit)
            {
                case "лет":
                    halfLife *= 365;
                    break;
                case "сут":
                    break;
                case "час":
                    halfLife /= 24;
                    break;
                case "мин":
                    halfLife /= 1440;
                    break;
                default:
                    return result;
            }
            if (halfLife < 60)
            {
                shortRad = rad;
                shortRadHalfLife = halfLife;
                valid = false;
                break;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = $"Период полураспада должен быть более 60 суток. Введенный вами радионуклид {shortRad} имеет период полураспада {shortRadHalfLife} суток. ОРИ на основе короткоживущих радионуклидов учитываются в формах 1.9 и 2.12."
            });
        }
        return result;
    }

    #endregion

    #region Check027

    //"Суммарная активность, Бк" положительное число, <= 10е+20 (колонка 8)
    private static List<CheckError> Check_027(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var activity = ConvertStringToExponential(forms[line].Activity_DB);
        if (string.IsNullOrEmpty(activity) || activity == "-") return result;
        if (!double.TryParse(activity,
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out var activityReal)
            || activity.Contains('-')
            || activityReal is <= 0 or > 10e+20)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = activity,
                Message = "Проверьте правильность предоставления сведений по активности."
            });
        }
        return result;
    }

    #endregion

    #region Check028

    private static List<CheckError> Check_028(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check029

    private static List<CheckError> Check_029(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check030

    private static List<CheckError> Check_030(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check031

    //Дата измерения активности <= дате операции
    private static List<CheckError> Check_031(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var operationDate = forms[line].OperationDate_DB;
        var activityMeasurementDate = forms[line].ActivityMeasurementDate_DB;
        var valid = DateOnly.TryParse(activityMeasurementDate, out var activityMeasurementDateReal)
                    && DateOnly.TryParse(operationDate, out var operationDateReal)
                    && activityMeasurementDateReal <= operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ActivityMeasurementDate_DB",
                Value = Convert.ToString(activityMeasurementDate),
                Message = "Дата измерения активности не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check032

    //Объем не пустая строка (колонка 10)
    private static List<CheckError> Check_032(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var volume = forms[line].Volume_DB;
        var valid = !string.IsNullOrWhiteSpace(volume);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Volume_DB",
                Value = Convert.ToString(volume),
                Message = "Графа не может быть пустой. Укажите оценочное значение объема ОРИ в круглых скобках."
            });
        }
        return result;
    }

    #endregion

    #region Check033

    //Объем >0, если в графе 12 значение 3 (колонка 10)
    private static List<CheckError> Check_033(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var volume = ConvertStringToExponential(forms[line].Volume_DB);
        var aggregateState = forms[line].AggregateState_DB;
        if (aggregateState is not 3) return result;
        if (!volume.Contains('e') && volume.Contains('+') ^ volume.Contains('-'))
        {
            volume = volume.Replace("+", "e+").Replace("-", "e-");
        }
        if (volume[0] == '(' && volume[^1] == ')')
        {
            volume = volume.Remove(volume.Length - 1, 1).Remove(0, 1);
        }
        var valid = double.TryParse(volume,
            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
            CultureInfo.CreateSpecificCulture("ru-RU"),
            out var volumeValue) && volumeValue > 0;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Volume_DB",
                Value = Convert.ToString(volume),
                Message = "Укажите объем ОРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check034

    //Масса не пустая строка (колонка 11)
    private static List<CheckError> Check_034(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var mass = forms[line].Mass_DB;
        var valid = !string.IsNullOrWhiteSpace(mass);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = Convert.ToString(mass),
                Message = "Графа не может быть пустой. Укажите оценочное значение массы ОРИ в круглых скобках."
            });
        }
        return result;
    }

    #endregion

    #region Check035

    //Масса >0, если в графе 12 значение 1 или 2 (колонка 11)
    private static List<CheckError> Check_035(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var mass = ConvertStringToExponential(forms[line].Mass_DB);
        if (string.IsNullOrEmpty(mass)) return result;
        var aggregateState = forms[line].AggregateState_DB;
        if (aggregateState is not 1 and not 2) return result;
        if (!mass.Contains('e') && mass.Contains('+') ^ mass.Contains('-'))
        {
            mass = mass.Replace("+", "e+").Replace("-", "e-");
        }
        if (mass[0] == '(' && mass[^1] == ')')
        {
            mass = mass.Remove(mass.Length - 1, 1).Remove(0, 1);
        }
        var valid = double.TryParse(mass,
            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
            CultureInfo.CreateSpecificCulture("ru-RU"),
            out var massValue) && massValue > 0;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = mass,
                Message = "Укажите положительную массу ОРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check036

    //Агрегатное состояния равно 1-3 (колонка 12)
    private static List<CheckError> Check_036(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var aggregateState = forms[line].AggregateState_DB;
        var valid = aggregateState is >= 1 and <= 3;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "AggregateState_DB",
                Value = Convert.ToString(aggregateState),
                Message = "Выберите значение для агрегатного состояния: 1 - твердое, 2 - жидкое, 3 - газообразное."
            });
        }
        return result;
    }

    #endregion

    #region Check037

    //Код формы собственности от 1 до 6, 9 (колонка 12)
    private static List<CheckError> Check_037(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is not (>= 1 and <= 6 or 9))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = Convert.ToString(propertyCode),
                Message = "Формат ввода данных не соответствует приказу. Выберите идентификатор, соответствующий форме собственности ЗРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check038

    // Если код формы собственности 2, необходимо указать примечание к ячейке
    private static List<CheckError> Check_038(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 13;
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is not 2) return result;
        var valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = Convert.ToString(propertyCode),
                Message = "Необходимо указать в примечании наименование субъекта Российской Федерации, в собственности которого находится объект учета."
            });
        }
        return result;
    }

    #endregion

    #region Check039

    // Если код формы собственности 3, необходимо указать примечание к ячейке
    private static List<CheckError> Check_039(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 13;
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is not 3) return result;
        var valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = Convert.ToString(propertyCode),
                Message = "Необходимо указать в примечании наименование муниципального образования, в собственности которого находится объект учета."
            });
        }
        return result;
    }

    #endregion

    #region Check040

    // Если код формы собственности 9, необходимо указать примечание к ячейке
    private static List<CheckError> Check_040(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 13;
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is not 9) return result;
        var valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = Convert.ToString(propertyCode),
                Message = "Необходимо указать в примечании наименование и адрес правообладателя (собственника или обладателя вещного права) объекта учета."
            });
        }
        return result;
    }

    #endregion

    #region Check041

    //8 или 14 чисел (колонка 14) если код формы собственности от 1 до 4 (колонка 13)
    private static List<CheckError> Check_041(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var owner = forms[line].Owner_DB;
        var propertyCode = forms[line].PropertyCode_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (propertyCode is < 1 or > 4) return result;
        var valid = owner != null && okpoRegex.IsMatch(owner);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = Convert.ToString(owner),
                Message = "Формат ввода данных не соответствует приказу. В случае, если правообладатель российское юридическое лицо, необходимо указать его код ОКПО."
            });
        }
        return result;
    }

    #endregion

    #region Check042

    //Правообладатель (колонка 14) из справочника ОКСМ, если код формы собственности 5 (колонка 13)
    private static List<CheckError> Check_042(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        if (propertyCode is not 5) return result;
        var valid = owner != null
                    && OKSM.Any(oksmEntry => oksmEntry["shortname"] == owner)
                    && !owner.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = Convert.ToString(owner),
                Message = "Формат ввода данных не соответствует приказу. В случае, если правообладатель иностранное государство, необходимо указать его краткое наименование в соответствии с ОКСМ."
            });
        }
        return result;
    }

    #endregion

    #region Check043

    //Правообладатель (колонка 14) из справочника ОКСМ и наличие примечания, если код формы собственности 6
    private static List<CheckError> Check_043(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 14;
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        if (propertyCode is not 6) return result;
        if (OKSM.All(oksmEntry => oksmEntry["shortname"] != owner)
            || owner.Equals("россия", StringComparison.CurrentCultureIgnoreCase))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = Convert.ToString(owner),
                Message = "В графе \"код ОКПО правообладателя\" необходимо указать краткое наименование страны в соответствии с ОКСМ."
            });
        }
        var valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = Convert.ToString(owner),
                Message = "Необходимо указать в примечании наименование и адрес правообладателя (собственника или обладателя иного вещного права) на ОРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check044

    //Вид документа от 1 до 15, 19 (колонка 15)
    private static List<CheckError> Check_044(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var documentVid = forms[line].DocumentVid_DB;
        var operationCode = forms[line].OperationCode_DB;
        var valid = documentVid is >= 1 and <= 15 or 19;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = Convert.ToString(documentVid),
                Message = "Формат ввода данных не соответствует приказу. Необходимо указать вид документа в соответствии с таблицей 3 приложения №2 к приказу Госкорпорации \"Росатом\" от 07.12.2020 №1/13-НПА."
            });
        }
        return result;
    }

    #endregion

    #region Check045

    //Номер документа не пустой (колонка 16)
    private static List<CheckError> Check_045(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var documentNumber = forms[line].DocumentNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(documentNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentNumber_DB",
                Value = Convert.ToString(documentNumber),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check046

    //Дата документа <= дате операции для всех кодов операции, кроме 10 и 41
    private static List<CheckError> Check_046(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        var documentDate = forms[line].DocumentDate_DB;
        if (operationCode is "10" or "41") return result;
        var valid = DateOnly.TryParse(documentDate, out var documentDateReal)
                    && DateOnly.TryParse(operationDate, out var operationDateReal)
                    && documentDateReal <= operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(documentDate),
                Message = "Дата документа не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check047

    //Дата документа = дате операции, если код операции 41
    private static List<CheckError> Check_047(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var documentDate = forms[line].DocumentDate_DB;
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        if (operationCode is not "41") return result;
        var valid = DateOnly.TryParse(documentDate, out var documentDateReal)
                    && DateOnly.TryParse(operationDate, out var operationDateReal)
                    && documentDateReal == operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(documentDate),
                Message = "Дата документа должна соответствовать дате операции."
            });
        }
        return result;
    }

    #endregion

    #region Check048

    //При коде операции 10, дата окончания ОП не позднее даты документа + 10 дней
    private static List<CheckError> Check_048(List<Form14> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10" };
        var documentDate = forms[line].DocumentDate_DB;

        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB) 
            || !(DateOnly.TryParse(rep.EndPeriod_DB, out var pEnd)
                && DateOnly.TryParse(documentDate, out var pMid))) return result;

        var valid = WorkdaysBetweenDates(pMid, pEnd) <= 10;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата окончания отчетного периода превышает дату акта инвентаризации более, чем на 10 рабочих дней."
            });
        }
        return result;
    }

    #endregion

    #region Check049

    //Код ОКПО поставщика/получателя равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_049(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "10", "11", "12", "15", "17", "18", "41", "42", "43", "46", "58", "61",
            "62", "65", "67", "68", "71", "72", "73", "74", "75", "97", "98", "99"
        };
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
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = "Для выбранного кода операции указывается код ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check050

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_050(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "25", "27", "28", "29", "35", "37", "38", "39", "63", "64"
        };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO)
                    && providerOrRecieverOKPO != repOKPO;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = "Для выбранного кода операции указывается код ОКПО контрагента."
            });
        }
        return result;
    }

    #endregion

    #region Check051

    //Код ОКПО поставщика/получателя состоит из 8/14 чисел или "минобороны" (колонка 18)
    private static List<CheckError> Check_051(List<Form14> forms, int line)
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
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = "Формат ввода данных не соответствует приказу. Следует указать код ОКПО контрагента, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check052

    //Код ОКПО поставщика/получателя из ОКСМ (не Россия), для определенных кодов операции, с примечанием (колонка 18)
    private static List<CheckError> Check_052(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 18;
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
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = "Необходимо добавить примечание."
            });
        }
        return result;
    }

    #endregion

    #region Check053

    //При определенных кодах операции, код ОКПО перевозчика равен "-" (колонка 19)
    private static List<CheckError> Check_053(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "10", "11", "12", "15", "17", "18", "41", "42", "43", "46", "53", "54",
            "58", "63", "64", "65", "66", "67", "68", "71", "72", "73", "74", "75", "97", "98"
        };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        var valid = transporterOKPO is "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = "При выбранном коде операции транспортирование не производится."
            });
        }
        return result;
    }

    #endregion

    #region Check054

    //При определенных кодах операции, код ОКПО перевозчика равен 8/14 цифр, допускается "прим." (колонка 19)
    private static List<CheckError> Check_054(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21", "25", "27", "28", "29", "31", "35", "37", "38", "39",
            "61", "62", "81", "82", "83", "84", "85", "86", "87", "88"
        };
        string[] transporterOKPOValid = { "прим.", "прим", "примечание", "примечания" };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        const byte graphNumber = 18;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (okpoRegex.IsMatch(transporterOKPO)) return result;

        var valid = okpoRegex.IsMatch(transporterOKPO)
                    || (transporterOKPOValid.Contains(transporterOKPO.ToLower())
                    && CheckNotePresence(new List<Form>(forms), notes, line, graphNumber));
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_13",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = "Необходимо указать код ОКПО организации перевозчика."
            });
        }
        return result;
    }

    #endregion

    #region Check055

    //Код ОКПО перевозчика состоит из 8/14 чисел или "минобороны" (колонка 19)
    private static List<CheckError> Check_055(List<Form14> forms, int line)
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
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = "Необходимо указать код ОКПО организации перевозчика, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check056

    //Не пустое поле (колонка 20)
    private static List<CheckError> Check_056(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var packName = forms[line].PackName_DB;
        var valid = !string.IsNullOrWhiteSpace(packName);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PackName_DB",
                Value = Convert.ToString(packName),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check057

    //Не пустое поле (колонка 21)
    private static List<CheckError> Check_057(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var packType = forms[line].PackType_DB;
        var valid = !string.IsNullOrWhiteSpace(packType);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PackType_DB",
                Value = Convert.ToString(packType),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check058

    //Не пустое поле (колонка 22)
    private static List<CheckError> Check_058(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var packNumber = forms[line].PackNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(packNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PackNumber_DB",
                Value = Convert.ToString(packNumber),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check061

    //Наличие строк дубликатов
    private static List<CheckError> Check_059(List<Form14> forms)
    {
        List<CheckError> result = new();
        HashSet<int> duplicatesLinesSet = new();
        var comparator = new CustomNullStringWithTrimComparer();
        for (var i = 0; i < forms.Count; i++)
        {
            var currentForm = forms[i];
            for (var j = i + 1; j < forms.Count;j++)
            {
                var formToCompare = forms[j];
                var isDuplicate = comparator.Compare(formToCompare.OperationCode_DB, currentForm.OperationCode_DB) == 0 
                                  && comparator.Compare(formToCompare.OperationDate_DB, currentForm.OperationDate_DB) == 0
                                  && comparator.Compare(formToCompare.PassportNumber_DB, currentForm.PassportNumber_DB) == 0
                                  && comparator.Compare(formToCompare.Name_DB, currentForm.Name_DB) == 0
                                  && formToCompare.Sort_DB == currentForm.Sort_DB
                                  && comparator.Compare(formToCompare.Radionuclids_DB, currentForm.Radionuclids_DB) == 0
                                  && comparator.Compare(formToCompare.Activity_DB, currentForm.Activity_DB) == 0
                                  && comparator.Compare(formToCompare.ActivityMeasurementDate_DB, currentForm.ActivityMeasurementDate_DB) == 0
                                  && comparator.Compare(formToCompare.Volume_DB, currentForm.Volume_DB) == 0
                                  && comparator.Compare(formToCompare.Mass_DB, currentForm.Mass_DB) == 0
                                  && formToCompare.AggregateState_DB == currentForm.AggregateState_DB
                                  && formToCompare.PropertyCode_DB == currentForm.PropertyCode_DB
                                  && comparator.Compare(formToCompare.Owner_DB, currentForm.Owner_DB) == 0
                                  && formToCompare.DocumentVid_DB == currentForm.DocumentVid_DB
                                  && comparator.Compare(formToCompare.DocumentNumber_DB, currentForm.DocumentNumber_DB) == 0
                                  && comparator.Compare(formToCompare.DocumentDate_DB, currentForm.DocumentDate_DB) == 0
                                  && comparator.Compare(formToCompare.ProviderOrRecieverOKPO_DB, currentForm.ProviderOrRecieverOKPO_DB) == 0;
                if (!isDuplicate) continue;
                duplicatesLinesSet.Add(i + 1);
                duplicatesLinesSet.Add(j + 1);
            }
        }
        var duplicateLines = string.Join(", ", duplicatesLinesSet.Order());
        if (duplicatesLinesSet.Count != 0)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = duplicateLines,
                Column = "2 - 17",
                Value = "",
                Message = $"Данные граф 2-18 в строках {duplicateLines} продублированы. Следует проверить правильность предоставления данных."
            });
        }
        return result;
    }

    #endregion

    #endregion

    #region GraphsList

    private static Dictionary<string, string> GraphsList { get; } = new()
    {
        { "NumberInOrder_DB", "01 - № п/п" },
        { "OperationCode_DB", "02 - Код операции" },
        { "OperationDate_DB", "03 - Дата операции" },
        { "PassportNumber_DB", "04 - Номер паспорта" },
        { "Name_DB", "05 - Наименование" },
        { "Sort_DB", "06 - Вид" },
        { "Radionuclids_DB", "07 - Радионуклиды" },
        { "Activity_DB", "08 - Активность, Бк." },
        { "ActivityMeasurementDate_DB", "09 - Дата измерения активности" },
        { "Volume_DB", "10 - Объем, куб.м." },
        { "Mass_DB", "11 - Масса, кг" },
        { "AggregateState_DB", "12 - Агрегатное состояние" },
        { "PropertyCode_DB", "13 - Код формы собственности" },
        { "Owner_DB", "14 - Код ОКПО правообладателя" },
        { "DocumentVid_DB", "15 - Вид документа" },
        { "DocumentNumber_DB", "16 - Номер документа" },
        { "DocumentDate_DB", "17 - Дата документа" },
        { "ProviderOrRecieverOKPO_DB", "18 - Код ОКПО поставщика или получателя" },
        { "TransporterOKPO_DB", "19 - Код ОКПО перевозчика" },
        { "PackName_DB", "20 - Наименование прибора, УКТ, упаковки" },
        { "PackType_DB", "21 - Тип прибора, УКТ, упаковки" },
        { "PackNumber_DB", "22 - Номер прибора, УКТ, упаковки" }
    };

    #endregion
}
