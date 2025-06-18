using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace WM_SalesOrderSplit
{
    class MeiktesEkptwseis
    {
        public static Company company = (Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
        //public static Company company = (Company)Menu.company;

        public static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form form = null;

            try
            {
                form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            }
            catch (Exception e)
            {
                Recordset rsError = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsError.DoQuery("INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit', '', '" + company.UserName + "', '" + e.Message.Replace("'", "") + " " + e.StackTrace.ToString().Replace("'", "") + "', 'ERROR')");
                return;
            }

            if ((
                pVal.FormType.ToString() == "133" ||
                pVal.FormType.ToString() == "139" ||
                pVal.FormType.ToString() == "140" ||
                pVal.FormType.ToString() == "149" ||
                pVal.FormType.ToString() == "179" ||
                pVal.FormType.ToString() == "180" ||
                pVal.FormType.ToString() == "234234567"
                )
                && pVal.ItemUID == "1"
                && pVal.BeforeAction == true
                && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                && (form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    || form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                )
            {

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Υπολογισμός Μεικτών Εκπτώσεων", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                try
                {
                    string sFinalItemCodes = "",
                       sCardCode = "",
                       sItem = "",
                       sDBDS = "",
                       sFormType = form.Type.ToString(); // "$[CURRENT_FORMTYPE]";

                    double dQty = 0.0;

                    Dictionary<string, double> ListItemCodes = new Dictionary<string, double>();

                    Recordset rsPrama = null;

                    List<string> lExistingItems = new List<string>();
                    List<string> lNonExistingItems = new List<string>();

                    switch (sFormType)
                    {
                        case "133":
                            sDBDS = "OINV";
                            break;
                        case "139":
                            sDBDS = "ORDR";
                            break;
                        case "140":
                            sDBDS = "ODLN";
                            break;
                        case "149":
                            sDBDS = "OQUT";
                            break;
                        case "179":
                            sDBDS = "ORIN";
                            break;
                        case "180":
                            sDBDS = "ORDN";
                            break;
                        case "234234567":
                            sDBDS = "ORRR";
                            break;
                        default:
                            return;
                    }

                    sCardCode = form.DataSources.DBDataSources.Item(sDBDS).GetValue("CardCode", 0).ToString();

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)form.Items.Item("38").Specific;

                    if (form.DataSources.DBDataSources.Item(sDBDS).GetValue("DocStatus", 0).ToString() == "C")
                    {
                        return;
                    }

                    for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                    {
                        sItem = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value.ToString();
                        dQty = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific).Value, System.Globalization.CultureInfo.InvariantCulture);

                        if (string.IsNullOrEmpty(sItem))
                        {
                            continue;
                        }

                        if (!ListItemCodes.ContainsKey(sItem))
                        {
                            ListItemCodes.Add(sItem, dQty);

                            sFinalItemCodes += ";" + sItem + "?" + dQty;
                        }
                        else
                        {
                            ListItemCodes[sItem] = ListItemCodes[sItem] + dQty;
                        }
                    }

                    sFinalItemCodes = sFinalItemCodes.Substring(1);

                    rsPrama = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    rsPrama.DoQuery("SELECT * FROM TKA_F_GET_MEIKTES_DISCOUNTS_V3('" + sFinalItemCodes + "', '" + sCardCode + "', CURRENT_DATE, '', '')");

                    if (rsPrama.RecordCount > 0)
                    {
                        rsPrama.MoveFirst();

                        while (!rsPrama.EoF)
                        { // vriskoume ta eidh set
                            if (!lExistingItems.Contains(rsPrama.Fields.Item(0).Value.ToString()))
                            {
                                lExistingItems.Add(rsPrama.Fields.Item(0).Value.ToString());
                            }

                            rsPrama.MoveNext();
                        }

                        // vriskoume ta eidh poy einai sthn paraggelia alla oxi se set
                        foreach (System.Collections.Generic.KeyValuePair<string, double> p in ListItemCodes)
                        {
                            if (!lExistingItems.Contains(p.Key))
                            {
                                lNonExistingItems.Add(p.Key);
                            }
                        }

                        // trwme ta eidh poy einai sthn paraggelia alla oxi se set
                        for (int i = 0; i < lNonExistingItems.Count; i++)
                        {
                            ListItemCodes.Remove(lNonExistingItems[i]);
                        }

                        double dTotalSetQty = 0.0;

                        // vriskoume to total twn eidwn set
                        foreach (KeyValuePair<string, double> p in ListItemCodes)
                        {
                            dTotalSetQty += p.Value;
                        }

                        rsPrama.DoQuery("SELECT * FROM TKA_F_GET_MEIKTES_DISCOUNTS_V3('" + sFinalItemCodes + "', '" + sCardCode + "', CURRENT_DATE, '', '')");

                        if (rsPrama.RecordCount > 0)
                        {
                            if (SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Βρέθηκαν τιμές σετ.\nΕνημέρωση γραμμών;", 2, "Ναι", "Όχι") != 1)
                            {
                                return;
                            }
                        }

                        rsPrama.MoveFirst();

                        form.Freeze(true);

                        while (!rsPrama.EoF)
                        {
                            for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                            {
                                sItem = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value.ToString();

                                if (rsPrama.Fields.Item("ITEMCODE").Value.ToString() == sItem)
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_TKA_Discount2").Cells.Item(i).Specific).Value = rsPrama.Fields.Item("DISCOUNT").Value.ToString();

                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_TKA_Discount1").Cells.Item(i).Specific).Value = string.Empty;
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_TKA_Discount3").Cells.Item(i).Specific).Value = string.Empty;

                                    continue;
                                }
                            }

                            rsPrama.MoveNext();
                        }
                    }
                }
                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace);
                }
                finally
                {
                    form.Freeze(false);
                }
            }
        }
    }
}
