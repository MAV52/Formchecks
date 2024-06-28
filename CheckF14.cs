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
        LoadDictionaries();
        var formsList = rep.Rows14.ToList<Form14>();
        errorList.AddRange(Check_002(rep));
        errorList.AddRange(Check_003(formsList, rep));
        errorList.AddRange(Check_004(formsList));
        foreach (var key in rep.Rows14)
        {
            var form = (Form14)key;
            var notes = rep.Notes.ToList<Note>();
            var forms10 = reps.Master_DB.Rows10.ToList<Form10>();
            errorList.AddRange(Check_001(formsList, currentFormLine));
            errorList.AddRange(Check_005(formsList, currentFormLine));
            errorList.AddRange(Check_006(formsList, currentFormLine));
            errorList.AddRange(Check_007(formsList, currentFormLine));
            errorList.AddRange(Check_008(formsList, notes, currentFormLine));
            errorList.AddRange(Check_009(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_010(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_011(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_012(formsList, currentFormLine));
            errorList.AddRange(Check_013(formsList, currentFormLine));
            errorList.AddRange(Check_014(formsList, currentFormLine));
            errorList.AddRange(Check_015(formsList, currentFormLine));
            errorList.AddRange(Check_016(formsList, currentFormLine));
            errorList.AddRange(Check_017(formsList, currentFormLine));
            errorList.AddRange(Check_018(formsList, currentFormLine));
            errorList.AddRange(Check_019(formsList, currentFormLine));
            errorList.AddRange(Check_020(formsList, currentFormLine));
            errorList.AddRange(Check_021(formsList, rep, currentFormLine));
            errorList.AddRange(Check_022(formsList, rep, currentFormLine));
            errorList.AddRange(Check_023(formsList, rep, currentFormLine));
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
            errorList.AddRange(Check_037(formsList, rep, currentFormLine));
            errorList.AddRange(Check_038(formsList, currentFormLine));
            errorList.AddRange(Check_039(formsList, currentFormLine));
            errorList.AddRange(Check_040(formsList, currentFormLine));
            errorList.AddRange(Check_041(formsList, currentFormLine));
            errorList.AddRange(Check_042(formsList, currentFormLine));
            errorList.AddRange(Check_043(formsList, currentFormLine));
            errorList.AddRange(Check_044(formsList, notes, currentFormLine));
            errorList.AddRange(Check_045(formsList, notes, currentFormLine));
            errorList.AddRange(Check_046(formsList, notes, currentFormLine));
            errorList.AddRange(Check_047(formsList, currentFormLine));
            errorList.AddRange(Check_048(formsList, currentFormLine));
            errorList.AddRange(Check_049(formsList, notes, currentFormLine));
            errorList.AddRange(Check_050(formsList, currentFormLine));
            errorList.AddRange(Check_051(formsList, currentFormLine));
            errorList.AddRange(Check_052(formsList, currentFormLine));
            errorList.AddRange(Check_053(formsList, currentFormLine));
            errorList.AddRange(Check_054(formsList, currentFormLine));
            errorList.AddRange(Check_055(formsList, rep, currentFormLine));
            errorList.AddRange(Check_056(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_057(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_058(formsList, currentFormLine));
            errorList.AddRange(Check_057(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_058(formsList, currentFormLine));
            errorList.AddRange(Check_059(formsList, notes, currentFormLine));
            errorList.AddRange(Check_060(formsList, currentFormLine));
            errorList.AddRange(Check_061(formsList, notes, currentFormLine));
            errorList.AddRange(Check_062(formsList, currentFormLine));
            errorList.AddRange(Check_063(formsList, currentFormLine));
            errorList.AddRange(Check_064(formsList, currentFormLine));
            errorList.AddRange(Check_065(formsList, currentFormLine));
            currentFormLine++;
        }
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

    //Заглушка
    private static List<CheckError> Check_001(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        return result;
    }

    #endregion

    #region Check002

    //Проверка правильности заполнения дат периода
    private static List<CheckError> Check_002(Report rep)
    {
        List<CheckError> result = new();
        if (!DateOnly.TryParse(rep.StartPeriod_DB, out var stPer))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = string.Empty,
                Column = string.Empty,
                Value = rep.StartPeriod_DB,
                Message = "Некорректно введена дата начала периода."
            });
        }
        if (!DateOnly.TryParse(rep.EndPeriod_DB, out var endPer))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = string.Empty,
                Column = string.Empty,
                Value = rep.EndPeriod_DB,
                Message = "Некорректно введена дата окончания периода."
            });
        }
        if (DateOnly.TryParse(rep.StartPeriod_DB, out _)
            && DateOnly.TryParse(rep.EndPeriod_DB, out _)
            && stPer > endPer)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = string.Empty,
                Column = string.Empty,
                Value = $"{rep.StartPeriod_DB} - {rep.EndPeriod_DB}",
                Message = "Дата начала периода превышает дату окончания периода."
            });
        }
        return result;
    }

    #endregion

    #region Check003

    //Проверка даты окончания периода
    private static List<CheckError> Check_003(List<Form14> forms14, Report rep)
    {
        var forms = forms14.Cast<Form1>().ToList();
        return CheckRepPeriod(forms, rep);
    }

    #endregion

    #region Check004

    //Наличие строк дубликатов (графы 2-18)
    private static List<CheckError> Check_004(List<Form14> forms)
    {
        List<CheckError> result = new();
        HashSet<int> duplicatesLinesSet = new();
        var comparator = new CustomNullStringWithTrimComparer();
        for (var i = 0; i < forms.Count; i++)
        {
            var currentForm = forms[i];
            for (var j = i + 1; j < forms.Count; j++)
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
                Column = "2 - 18",
                Value = "",
                Message = $"Данные граф 2-18 в строках {duplicateLines} продублированы. " +
                          $"Следует проверить правильность предоставления данных."
            });
        }
        return result;
    }

    #endregion

    #region Check005

    //Нумерация строк (графа 1)
    private static List<CheckError> Check_005(List<Form14> forms, int line)
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
                Message = "Номера строк должны располагаться по порядку, без пропусков или дублирования номеров."
            });
        }
        return result;
    }

    #endregion

    #region Check006

    //Код операции из списка (графа 2)
    private static List<CheckError> Check_006(List<Form14> forms, int line)
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

    #region Check007

    //Для кодов операции 12, 42, в радионуклидах должен быть указан хоть один из списка (графа 6)
    private static List<CheckError> Check_007(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var radionuclid = forms[line].Radionuclids_DB ?? string.Empty;
        string[] applicableOperationCodes = { "12", "42" };
        var radionuclidValid = new[]
        {
            "плутоний-238", "плутоний-239", "плутоний-240", "уран-233", "уран-235", "уран-238", "нептуний-237",
            "америций-241", "америций 243", "калифорний-252", "торий-232", "литий-6", "дейтерий", "тритий"
        };
        if (string.IsNullOrWhiteSpace(radionuclid) || !applicableOperationCodes.Contains(operationCode)) return result;
        var valid = radionuclidValid.Any(nuclid => radionuclid.ToLower().Contains(nuclid));
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclid),
                Message = "В графе 6 не представлены сведения о радионуклидах, которые могут быть отнесены к ЯМ. " +
                          "Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check008

    //Для определенных кодов операции должно быть примечание (графа 2)
    private static List<CheckError> Check_008(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 2;
        var operationCode = forms[line].OperationCode_DB;
        string[] applicableOperationCodes = { "29", "39", "97", "98", "99" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = CheckNotePresence(notes, line, graphNumber);
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

    #region Check009

    //Код ОКПО поставщика/получателя равен коду ОКПО отчитывающейся организации (графа 18)
    private static List<CheckError> Check_009(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "53" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB ?? string.Empty;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = providerOrRecieverOkpo.Equals(repOkpo);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "В графе 18 необходимо указать код ОКПО своей организации. В случае, " +
                          "если зарядку/разрядку осуществляла подрядная организация, следует использовать код операции 54."
            });
        }
        return result;
    }

    #endregion

    #region Check010

    //Код ОКПО правообладателя (графа 14) равен коду ОКПО юридического лица или обособленного подразделения
    private static List<CheckError> Check_010(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "11", "12", "15", "28", "38", "41", "48", "63", "64", "65", "73", "81", "85", "88" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var owner = forms[line].Owner_DB;
        var okpoRepJur = forms10[0].Okpo_DB ?? string.Empty;
        var okpoRepTerPodr = forms10[1].Okpo_DB ?? string.Empty;
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

    #region Check011

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации (графа 18)
    private static List<CheckError> Check_011(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "54" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB ?? string.Empty;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = !string.IsNullOrWhiteSpace(providerOrRecieverOkpo) 
                    && !providerOrRecieverOkpo.Equals(repOkpo) 
                    && OkpoRegex.IsMatch(providerOrRecieverOkpo);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "В графе 18 необходимо указать ОКПО подрядной организации."
            });
        }
        return result;
    }

    #endregion

    #region Check012

    //TODO
    //Проверка на соответствие СНК
    private static List<CheckError> Check_012(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check013

    //TODO
    //Наличие в формах 1.5/1.6 кода 44
    private static List<CheckError> Check_013(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check014

    //TODO
    //Наличие кода 35
    private static List<CheckError> Check_014(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check015

    //TODO
    //Проверка наличия учётной единицы
    private static List<CheckError> Check_015(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check016

    //TODO
    //Наличие кода 27
    private static List<CheckError> Check_016(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check017

    //TODO
    //наличие учётной единицы в форме 1.6
    private static List<CheckError> Check_017(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check018

    //TODO
    //Наличие кода 46
    private static List<CheckError> Check_018(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check019

    //TODO
    //Наличие кода 61
    private static List<CheckError> Check_019(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check020

    //TODO
    //Наличие кода 65 в форме 1.1
    private static List<CheckError> Check_020(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check021

    //Дата операции корректно заполнена (графа 3)
    private static List<CheckError> Check_021(List<Form14> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var opDate = forms[line].OperationDate_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(opDate) || opDate is "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(opDate),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        else if (!DateOnly.TryParse(opDate, out _))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(opDate),
                Message = "Формат ввода данных не соответствует приказу. Некорректно заполнена дата операции."
            });
        }
        return result;
    }

    #endregion

    #region Check022

    //Дата документа входит в отчетный период с учетом срока подачи отчета в днях (графа 3)
    private static List<CheckError> Check_022(List<Form14> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB;
        if (opCode is "10"
            || !DateOnly.TryParse(rep.StartPeriod_DB, out var pStart)
            || !DateOnly.TryParse(rep.EndPeriod_DB, out var pEnd)
            || !DateOnly.TryParse(forms[line].OperationDate_DB, out var opDate))
        {
            return result;
        }
        var valid = opDate >= pStart && opDate <= pEnd;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(forms[line].OperationDate_DB),
                Message = "Дата операции не входит в отчетный период."
            });
        }
        return result;
    }

    #endregion

    #region Check023

    //При коде операции 10, дата документа должна попадать в отчетный период
    private static List<CheckError> Check_023(List<Form14> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB;
        if (opCode is not "10"
            || !DateOnly.TryParse(forms[line].DocumentDate_DB, out var docDate)
            || !DateOnly.TryParse(rep.StartPeriod_DB, out var stDate)
            || !DateOnly.TryParse(rep.EndPeriod_DB, out var endDate))
        {
            return result;
        }
        var valid = docDate >= stDate && docDate <= endDate;
        if(!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(forms[line].DocumentDate_DB),
                Message = "Дата документа не входит в отчетный период. Для операции инвентаризации, " +
                          "срок предоставления отчета исчисляется с даты утверждения акта инвентаризации."
            });
        }
        return result;
    }

    #endregion

    #region Check024

    //Номер паспорта не пустая строка (графа 4)
    private static List<CheckError> Check_024(List<Form14> forms, int line)
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
                Message = "Графа не может быть пустой. Укажите номера паспорта (сертификата). " +
                          "В случае, если номер отсутствует укажите символ \"-\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check025

    //Наименование не пустая строка (графа 5)
    private static List<CheckError> Check_025(List<Form14> forms, int line)
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
                Message = "Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check026

    //Вид от 4 до 12 (графа 6)
    private static List<CheckError> Check_026(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var sort = forms[line].Sort_DB ?? 0;
        var valid = sort is >= 4 and <= 12;
        if (!valid) 
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Sort_DB",
                Value = Convert.ToString(sort),
                Message = $"Формат ввода данных не соответствует приказу. Необходимо указать вид РВ " +
                          $"в соответствии с таблицей 4 приложения №2 к приказу Госкорпорации \"Росатом\" от 07.12.2020 №1/13-НПА."
            });
        }
        return result;
    }

    #endregion

    #region Check027

    //Агрегатное состояния 1, если вид 7 (графа 12)
    private static List<CheckError> Check_027(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var sort = forms[line].Sort_DB ?? 0;
        var agrState = forms[line].AggregateState_DB ?? 0;
        if (sort != 7) return result;
        var valid = agrState is 1;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "AggregateState_DB",
                Value = Convert.ToString(agrState),
                Message = $"Для выбранного вида ОРИ ({sort}) - агрегатное состояние (графа 12) - жидкое (значение 1)."
            });
        }
        return result;
    }

    #endregion

    #region Check028

    //Агрегатное состояния 2, если вид 4-6 (графа 12)
    private static List<CheckError> Check_028(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var sort = forms[line].Sort_DB ?? 0;
        var agrState = forms[line].AggregateState_DB ?? 0;
        if (sort is not (4 or 5 or 6)) return result;
        var valid = agrState is 2;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "AggregateState_DB",
                Value = Convert.ToString(agrState),
                Message = $"Для выбранного вида ОРИ ({sort}) - агрегатное состояние (графа 12) - твёрдое (значение 2)."
            });
        }
        return result;
    }

    #endregion

    #region Check029

    //Агрегатное состояния 1 или 2, если вид 4-7 (графа 12)
    private static List<CheckError> Check_029(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var sort = forms[line].Sort_DB ?? 0;
        var agrState = forms[line].AggregateState_DB ?? 0;
        if (sort is not (4 or 5 or 6 or 7)) return result;
        var valid = agrState is 1 or 2;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "AggregateState_DB",
                Value = Convert.ToString(agrState),
                Message = $"Для выбранного вида ОРИ ({sort}) - агрегатное состояние не может быть газообразным."
            });
        }
        return result;
    }

    #endregion

    #region Check030

    //Радионуклиды заполнены (графа 7)
    private static List<CheckError> Check_030(List<Form14> forms, int line)
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

    #region Check031

    //Проверка радионуклидов по справочнику (графа 7)
    private static List<CheckError> Check_031(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var rads = forms[line].Radionuclids_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(rads) || rads.Trim() == "-") return result;
        var radsSet = rads
            .ToLower()
            .Replace(" ", string.Empty)
            .Replace(',', ';')
            .Split(';')
            .ToHashSet();
        if (radsSet.Count == 1 && R.All(phEntry => phEntry["name"] != radsSet.First()))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = rads,
                Message = "Формат ввода данных не соответствует приказу. " +
                          "Наименование радионуклида указывается названием химического элемента на русском языке " +
                          "с указанием через дефис массового числа изотопа."
            });
            return result;
        }
        if (rads.Contains(',') 
            || rads.Contains('+') 
            || !radsSet
                .All(rad => R
                    .Any(phEntry => phEntry["name"] == rad)))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = rads,
                Message = "Формат ввода данных не соответствует приказу. " +
                          "Наименование радионуклида указывается названием химического элемента на русском языке " +
                          "с указанием через дефис массового числа изотопа, радионуклиды перечисляются через точку с запятой."
            });
        }
        return result;
    }

    #endregion

    #region Check032

    //Каждый из радионуклидов не короткоживущий, >=60 суток (графа 7)
    private static List<CheckError> Check_032(List<Form14> forms, int line)
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
                Message = $"Период полураспада должен быть более 60 суток. " +
                          $"Введенный вами радионуклид {shortRad} имеет период полураспада {shortRadHalfLife} суток. " +
                          $"ОРИ на основе короткоживущих радионуклидов учитываются в формах 1.9 и 2.12."
            });
        }
        return result;
    }

    #endregion

    #region Check033

    //"Суммарная активность, Бк" > 0, <= 10е+20 (графа 8)
    private static List<CheckError> Check_033(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var activity = ConvertStringToExponential(forms[line].Activity_DB);
        if (string.IsNullOrEmpty(activity) || activity == "-") return result;
        if (!double.TryParse(activity,
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out var activityReal)
            || activity.Contains('-'))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = Convert.ToString(forms[line].Activity_DB),
                Message = "Проверьте правильность предоставления сведений по активности."
            });
        }
        else if (activityReal <= 0)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = Convert.ToString(forms[line].Activity_DB),
                Message = "Суммарная активность должна быть более нуля. Проверьте правильность введённых данных."
            });
        }
        else if (activityReal > 10e+20)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = Convert.ToString(forms[line].Activity_DB),
                Message = "Указано слишком большое значение суммарной активности. Проверьте правильность введённых данных."
            });
        }
        return result;
    }

    #endregion

    #region Check034

    //TODO
    //Сверка МЗА для одного радионуклида
    private static List<CheckError> Check_034(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check035

    //TODO
    //Сверка МЗА для нескольких радионуклидов
    private static List<CheckError> Check_035(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check036

    //TODO
    //Сверка МЗА для равновесных радионуклидов
    private static List<CheckError> Check_036(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        return result;
    }

    #endregion

    #region Check037

    //Дата измерения активности корректно заполнена (графа 9)
    private static List<CheckError> Check_037(List<Form14> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var activityMeasurementDate = forms[line].ActivityMeasurementDate_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(activityMeasurementDate) || activityMeasurementDate is "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ActivityMeasurementDate_DB",
                Value = Convert.ToString(activityMeasurementDate),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        else if (!DateOnly.TryParse(activityMeasurementDate, out _) 
                 && activityMeasurementDate.ToLower() is not ("прим." or "прим" or "примечание" or "примечания"))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ActivityMeasurementDate_DB",
                Value = Convert.ToString(activityMeasurementDate),
                Message = "Формат ввода данных не соответствует приказу. Некорректно заполнена дата измерения активности."
            });
        }
        return result;
    }

    #endregion

    #region Check038

    //Дата измерения активности (графа 9) <= дате операции
    private static List<CheckError> Check_038(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        if (!DateOnly.TryParse(forms[line].ActivityMeasurementDate_DB, out var activityMeasurementDate)
            || !DateOnly.TryParse(forms[line].OperationDate_DB, out var opDate))
        {
            return result;
        }
        var valid = activityMeasurementDate <= opDate;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ActivityMeasurementDate_DB",
                Value = Convert.ToString(forms[line].ActivityMeasurementDate_DB),
                Message = "Дата измерения активности не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check039

    //Объем не пустая строка, больше 0 (графа 10)
    private static List<CheckError> Check_039(List<Form14> forms, int line)
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
                Message = "Графа не может быть пустой. При отсутствии сведений в круглых скобках указывается оценочное значение."
            });
        }
        return result;
    }

    #endregion

    #region Check040

    //Объем > 0, если в графе 12 значение 3 (графа 10)
    private static List<CheckError> Check_040(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var volume = ConvertStringToExponential(forms[line].Volume_DB);
        var aggregateState = forms[line].AggregateState_DB;
        if (aggregateState is not 3) return result;
        var valid = double.TryParse(volume,
            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
            CultureInfo.CreateSpecificCulture("ru-RU"),
            out var volumeValue) 
                    && volumeValue > 0;
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

    #region Check041

    //Масса не пустая строка (графа 11)
    private static List<CheckError> Check_041(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var mass = ConvertStringToExponential(forms[line].Mass_DB);
        var agrState = forms[line].AggregateState_DB ?? 0;
        if (agrState is not (1 or 2)) return result;
        if (string.IsNullOrWhiteSpace(mass) || mass is "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = Convert.ToString(mass),
                Message = "Графа не может быть пустой. При отсутствии сведений в круглых скобках указывается оценочное значение."
            });
        }
        else if (!double.TryParse(mass,
                     NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
                     CultureInfo.CreateSpecificCulture("ru-RU"),
                     out var massValue))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = Convert.ToString(mass),
                Message = "Некорректно введено значение массы."
            });
        }
        else if (massValue <= 0)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = Convert.ToString(mass),
                Message = "Значение массы должно быть более нуля."
            });
        }
        return result;
    }

    #endregion

    #region Check042

    //Агрегатное состояния равно 1-3 (графа 12)
    private static List<CheckError> Check_042(List<Form14> forms, int line)
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

    #region Check043

    //Код формы собственности от 1 до 6, 9 (графа 12)
    private static List<CheckError> Check_043(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        var valid = propertyCode is >= 1 and <= 6 or 9;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = Convert.ToString(propertyCode),
                Message = "Формат ввода данных не соответствует приказу. " +
                          "Выберите идентификатор, соответствующий форме собственности ЗРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check044

    // Если код формы собственности 2, необходимо указать примечание к ячейке (графа 13)
    private static List<CheckError> Check_044(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 13;
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is not 2) return result;
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = Convert.ToString(propertyCode),
                Message = "Необходимо указать в примечании наименование субъекта Российской Федерации, " +
                          "в собственности которого находится объект учета."
            });
        }
        return result;
    }

    #endregion

    #region Check045

    // Если код формы собственности 3, необходимо указать примечание к ячейке (графа 13)
    private static List<CheckError> Check_045(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 13;
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is not 3) return result;
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = Convert.ToString(propertyCode),
                Message = "Необходимо указать в примечании наименование муниципального образования, " +
                          "в собственности которого находится объект учета."
            });
        }
        return result;
    }

    #endregion

    #region Check046

    // Если код формы собственности 9, необходимо указать примечание к ячейке (графа 13)
    private static List<CheckError> Check_046(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 13;
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is not 9) return result;
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = Convert.ToString(propertyCode),
                Message = "Необходимо указать в примечании наименование и адрес правообладателя " +
                          "(собственника или обладателя вещного права) объекта учета."
            });
        }
        return result;
    }

    #endregion

    #region Check047

    //8 или 14 чисел (графа 14) если код формы собственности от 1 до 4 (графа 13)
    private static List<CheckError> Check_047(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var owner = forms[line].Owner_DB ?? string.Empty;
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is < 1 or > 4) return result;
        var valid = OkpoRegex.IsMatch(owner);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = Convert.ToString(owner),
                Message = "Формат ввода данных не соответствует приказу. В случае, " +
                          "если правообладатель российское юридическое лицо, необходимо указать его код ОКПО."
            });
        }
        return result;
    }

    #endregion

    #region Check048

    //Правообладатель (графа 14) из справочника ОКСМ, если код формы собственности 5 (графа 13)
    private static List<CheckError> Check_048(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB ?? string.Empty;
        if (propertyCode is not 5) return result;
        var valid = OKSM.Any(oksmEntry => oksmEntry["shortname"] == owner) 
                    && !owner.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = Convert.ToString(owner),
                Message = "Формат ввода данных не соответствует приказу. В случае, если правообладатель иностранное государство, " +
                          "необходимо указать его краткое наименование в соответствии с ОКСМ."
            });
        }
        return result;
    }

    #endregion

    #region Check049

    //Правообладатель (графа 14) из справочника ОКСМ и наличие примечания, если код формы собственности 6
    private static List<CheckError> Check_049(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 14;
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB ?? string.Empty;
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
                Message = "В графе \"код ОКПО правообладателя\", необходимо указать краткое наименование страны в соответствии с ОКСМ."
            });
        }
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = Convert.ToString(owner),
                Message = "Необходимо указать в примечании наименование и адрес правообладателя " +
                          "(собственника или обладателя иного вещного права) на ОРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check050

    //Вид документа от 1 до 15, 19 (графа 15)
    private static List<CheckError> Check_050(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var documentVid = forms[line].DocumentVid_DB;
        var valid = documentVid is >= 1 and <= 15 or 19;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = Convert.ToString(documentVid),
                Message = "Формат ввода данных не соответствует приказу. Необходимо указать вид документа в соответствии с таблицей 3 " +
                          "приложения №2 к приказу Госкорпорации \"Росатом\" от 07.12.2020 №1/13-НПА."
            });
        }
        return result;
    }

    #endregion

    #region Check051

    //Номер документа не пустой (графа 16)
    private static List<CheckError> Check_051(List<Form14> forms, int line)
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

    #region Check052

    //Дата документа корректно заполнена (графа 17)
    private static List<CheckError> Check_052(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var docDate = forms[line].DocumentDate_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(docDate) || docDate is "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(docDate),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        else if (!DateOnly.TryParse(docDate, out _))
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(docDate),
                Message = "Формат ввода данных не соответствует приказу. Некорректно заполнена дата документа."
            });
        }
        return result;
    }

    #endregion

    #region Check053

    //Дата документа <= дате операции для всех кодов операции, кроме 10 и 41
    private static List<CheckError> Check_053(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB ?? string.Empty;
        var documentDate = forms[line].DocumentDate_DB ?? string.Empty;
        if (operationCode is "10" or "41"
            || !DateOnly.TryParse(documentDate, out var documentDateReal)
            || !DateOnly.TryParse(operationDate, out var operationDateReal))
        {
            return result;
        }
        var valid = documentDateReal <= operationDateReal;
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

    #region Check054

    //Дата документа = дате операции, если код операции 41
    private static List<CheckError> Check_054(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var documentDate = forms[line].DocumentDate_DB;
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        if (operationCode is not "41"
            || !DateOnly.TryParse(documentDate, out var documentDateReal)
            || !DateOnly.TryParse(operationDate, out var operationDateReal))
        {
            return result;
        }
        var valid = documentDateReal == operationDateReal;
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

    #region Check055

    //При коде операции 10, дата окончания ОП не позднее даты документа + 10 дней
    private static List<CheckError> Check_055(List<Form14> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        var documentDate = forms[line].DocumentDate_DB ?? string.Empty;
        var endPeriod = rep.EndPeriod_DB ?? string.Empty;

        if (opCode is not "10"
            || !DateOnly.TryParse(endPeriod, out var pEnd) 
            || !DateOnly.TryParse(documentDate, out var docDate)) return result;

        var valid = WorkdaysBetweenDates(docDate, pEnd) <= 10;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата документа не входит в отчетный период. Для операции инвентаризации, " +
                          "срок предоставления отчета исчисляется с даты утверждения акта инвентаризации."
            });
        }
        return result;
    }

    #endregion

    #region Check056

    //Код ОКПО поставщика/получателя равен коду ОКПО отчитывающейся организации + 8/14 цифр (графа 18)
    private static List<CheckError> Check_056(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "10", "11", "12", "15", "17", "18", "41", "42", "43", "46", "58", "61",
            "62", "65", "67", "68", "71", "72", "73", "74", "75", "97", "98", "99"
        };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(providerOrRecieverOkpo)
                    && providerOrRecieverOkpo == repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Для выбранного кода операции указывается код ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check057

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (графа 18)
    private static List<CheckError> Check_057(List<Form14> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "25", "27", "28", "29", "35", "37", "38", "39", "63", "64" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = OkpoRegex.IsMatch(providerOrRecieverOkpo)
                    && providerOrRecieverOkpo != repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Для выбранного кода операции указывается код ОКПО контрагента."
            });
        }
        return result;
    }

    #endregion

    #region Check058

    //Код ОКПО поставщика/получателя состоит из 8/14 чисел или "минобороны" (колонка 18)
    private static List<CheckError> Check_058(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        if (operationCode is not ("22" or "32")) return result;
        var valid = OkpoRegex.IsMatch(providerOrRecieverOkpo)
                    || providerOrRecieverOkpo.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Формат ввода данных не соответствует приказу. Следует указать код ОКПО контрагента, " +
                          "либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check059

    //Код ОКПО поставщика/получателя из ОКСМ (не Россия), для определенных кодов операции, с примечанием (графа 18)
    private static List<CheckError> Check_059(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 18;
        string[] applicableOperationCodes = { "81", "82", "83", "84", "85", "86", "87", "88" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OKSM.Any(oksmEntry => oksmEntry["shortname"] == providerOrRecieverOkpo)
                    && !providerOrRecieverOkpo.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Необходимо добавить примечание."
            });
        }
        return result;
    }

    #endregion

    #region Check060

    //При определенных кодах операции, код ОКПО перевозчика равен "-" (графа 19)
    private static List<CheckError> Check_060(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "10", "11", "12", "15", "17", "18", "41", "42", "43", "46", "53", "54", "58", 
            "63", "64", "65", "66", "67", "68", "71", "72", "73", "74", "75", "97", "98"
        };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var transporterOkpo = forms[line].TransporterOKPO_DB;
        var valid = transporterOkpo is "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOkpo),
                Message = "При выбранном коде операции транспортирование не производится. Укажите символ \"-\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check061

    //Код ОКПО перевозчика равен 8/14 цифр, допускается "прим." (графа 19)
    private static List<CheckError> Check_061(List<Form14> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21", "25", "27", "28", "29", "31", "35", "37", "38", "39",
            "61", "62", "81", "82", "83", "84", "85", "86", "87", "88"
        };
        string[] transporterOkpoValid = { "прим.", "прим", "примечание", "примечания" };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOkpo = forms[line].TransporterOKPO_DB;
        const byte graphNumber = 18;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(transporterOkpo)
                    || (transporterOkpoValid.Contains(transporterOkpo.ToLower())
                    && CheckNotePresence(notes, line, graphNumber));
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOkpo),
                Message = "Необходимо указать код ОКПО организации перевозчика."
            });
        }
        return result;
    }

    #endregion

    #region Check062

    //Код ОКПО перевозчика состоит из 8/14 чисел или "минобороны" (графа 19)
    private static List<CheckError> Check_062(List<Form14> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "22", "32" };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOkpo = forms[line].TransporterOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(transporterOkpo)
                    || transporterOkpo.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_14",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOkpo),
                Message = "Необходимо указать код ОКПО организации перевозчика, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check063

    //Наименование прибора заполнено (графа 20)
    private static List<CheckError> Check_063(List<Form14> forms, int line)
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

    #region Check064

    //Тип прибора заполнен (графа 21)
    private static List<CheckError> Check_064(List<Form14> forms, int line)
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

    #region Check065

    //Номер прибора заполнен (графа 22)
    private static List<CheckError> Check_065(List<Form14> forms, int line)
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
