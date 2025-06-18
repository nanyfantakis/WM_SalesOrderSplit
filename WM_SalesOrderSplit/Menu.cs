//using SAPbouiCOM;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using WM_SalesOrderSplit.Prama;

namespace WM_SalesOrderSplit
{
    class Menu
    {
        public static Company company = (Company)Application.SBO_Application.Company.GetDICompany();

        public static SAPbouiCOM.Form form = null;

        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus;
            SAPbouiCOM.MenuItem oMenuItem;

            SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
            oMenuItem = Application.SBO_Application.Menus.Item("2048"); // ar

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oCreationPackage.UniqueID = "SalesOrderSplitMenu";
            oCreationPackage.String = "Ειδικός Διαχωρισμός Παραγγελιών";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = 2;

            oMenus = oMenuItem.SubMenus;

            try
            {
                oMenus.AddEx(oCreationPackage);

                Application.SBO_Application.StatusBar.SetText("Ειδικός Διαχωρισμός Παραγγελιών Add-On Connected Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception)
            {
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "SalesOrderSplitMenu")
                {
                    for (int i = 0; i< Application.SBO_Application.Forms.Count; i++)
                    {
                        if (Application.SBO_Application.Forms.Item(i).TypeEx == "WMSalesOrderSplit")
                        {
                            Application.SBO_Application.MessageBox("Η Φόρμα Ειδικού Διαχωρισμού Παραγγελιών είναι ήδη ανοιχτή.");
                            return;
                        }
                    }
                    
                    string sFormAsXML = Stringia.sWM_SalesOrderSplit;

                    SAPbouiCOM.FormCreationParams oCreationParams = null;

                    oCreationParams = (SAPbouiCOM.FormCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    oCreationParams.XmlData = sFormAsXML;

                    form = Application.SBO_Application.Forms.AddEx(oCreationParams);

                    Recordset rsGetPrama = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    form.DataSources.UserDataSources.Item("SlctAllPn1").Value = "Y";
                    form.DataSources.UserDataSources.Item("ElleipsiDS").Value = "N";
                    form.DataSources.UserDataSources.Item("SeriesTo").Value = "8";
                    form.DataSources.UserDataSources.Item("ShipTypeDS").Value = "-1";

                    rsGetPrama = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    rsGetPrama.DoQuery("SELECT \"Series\", \"SeriesName\" FROM NNM1 WHERE \"ObjectCode\" = '17' AND \"SeriesName\" != 'ΕΠΙΛΕΞΤΕ'");

                    SAPbouiCOM.ComboBox oCB1 = (SAPbouiCOM.ComboBox)form.Items.Item("SeriesTo").Specific;

                    while (!rsGetPrama.EoF)
                    {
                        oCB1.ValidValues.Add(rsGetPrama.Fields.Item(0).Value.ToString(), rsGetPrama.Fields.Item(1).Value.ToString());

                        rsGetPrama.MoveNext();
                    }

                    rsGetPrama.DoQuery("SELECT T0.\"Name\", T0.\"Code\" FROM \"@TKA_SUBCATEGORY\" T0 " +
                                        " UNION ALL SELECT '', '' FROM DUMMY ORDER BY 1");

                    oCB1 = (SAPbouiCOM.ComboBox)form.Items.Item("SubCatCB").Specific;

                    while (!rsGetPrama.EoF)
                    {
                        oCB1.ValidValues.Add(rsGetPrama.Fields.Item(0).Value.ToString(), rsGetPrama.Fields.Item(1).Value.ToString());

                        rsGetPrama.MoveNext();
                    }

                    /*rsGetPrama.DoQuery("SELECT T0.\"TrnspCode\", T0.\"TrnspName\" FROM OSHP T0 " +
                                        " UNION ALL SELECT '-1', '' FROM DUMMY ORDER BY 1");*/
                    rsGetPrama.DoQuery("SELECT \"TrnspCode\", N'Ναι' FROM OSHP WHERE \"TrnspCode\" = 4  UNION ALL SELECT -1, N'Όχι' FROM DUMMY");

                    oCB1 = (SAPbouiCOM.ComboBox)form.Items.Item("ShipTypeCB").Specific;

                    while (!rsGetPrama.EoF)
                    {
                        oCB1.ValidValues.Add(rsGetPrama.Fields.Item(0).Value.ToString(), rsGetPrama.Fields.Item(1).Value.ToString());

                        rsGetPrama.MoveNext();
                    }

                    ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).DataTable.ExecuteQuery("SELECT 'N' \"Check\", \"Series\", \"SeriesName\", \"Remark\" FROM NNM1 WHERE \"ObjectCode\" = '23' AND \"SeriesName\" != 'ΕΠΙΛΕΞΤΕ'");

                    ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).Columns.Item("Series").Editable = false;
                    ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).Columns.Item("SeriesName").Editable = false;
                    ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).Columns.Item("Remark").Editable = false;

                    ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

                    ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).AutoResizeColumns();

                    // StaticText Stoxou
                    form.Items.Item("Item_5").TextStyle = 1;
                    form.Items.Item("Item_9").TextStyle = 1;

                    form.Items.Item("CardCodeED").Height = 20;
                    form.Items.Item("DocDateEDF").Height = 20;
                    form.Items.Item("DocDateEDT").Height = 20;
                    form.Items.Item("TaxDateEDF").Height = 20;
                    form.Items.Item("TaxDateEDT").Height = 20;
                    form.Items.Item("SubCatCB").Height = 20;
                    form.Items.Item("ElleipsiCB").Height = 20;

                    form.Items.Item("DocDateED").Height = 20;
                    form.Items.Item("SeriesTo").Height = 20;

                    form.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMMdd");
                    form.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");
                    form.DataSources.UserDataSources.Item("DocDateDS").Value = DateTime.Now.ToString("yyyyMMdd");

                    form.Visible = true;

                    Initialize_Form(ref form);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        public void Initialize_Form(ref SAPbouiCOM.Form form)
        {
            try
            {
                //form = Application.SBO_Application.Forms.ActiveForm;

                Users oUser = (Users)company.GetBusinessObject(BoObjectTypes.oUsers);

                oUser.GetByKey(((ICompany)Application.SBO_Application.Company.GetDICompany()).UserSignature);

                form.DataSources.UserDataSources.Item("SlctAllPn1").Value = "Y";

                ((SAPbouiCOM.Button)form.Items.Item("SelectBTN").Specific).ClickAfter += SelectBTN_ClickAfter;
                ((SAPbouiCOM.Button)form.Items.Item("BackBtn").Specific).ClickAfter += BackBtn_ClickAfter;
                ((SAPbouiCOM.Button)form.Items.Item("NextBtn").Specific).ClickAfter += NextBtn_ClickAfter;
                //((SAPbouiCOM.Button)form.Items.Item("ExecBtn").Specific).ClickAfter += ExecBtn_ClickAfter;
                ((SAPbouiCOM.Button)form.Items.Item("RefreshBT").Specific).ClickAfter += NextBtn_ClickAfter;

                ((SAPbouiCOM.EditText)form.Items.Item("CardCodeED").Specific).ValidateAfter += CardCodeED_ValidateAfter;
                ((SAPbouiCOM.EditText)form.Items.Item("CardCodeED").Specific).ChooseFromListAfter += CardCodeED_ChooseFromListAfter;
            }
            catch (Exception e)
            {
                Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace);
            }
        }

        private void SelectBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)form.Items.Item("ResultsGRD").Specific;

            form.Freeze(true);

            try
            {
                if (form.DataSources.UserDataSources.Item("SlctAllPn1").Value == "Y")
                {
                    for (int i = 0; i < oGrid.Rows.Count; i++)
                    {
                        if (oGrid.DataTable.GetValue("ZERO_IS_INVALID", i).ToString() != "0")
                        {
                            oGrid.DataTable.SetValue("Check", oGrid.GetDataTableRowIndex(i), "Y");
                        }
                    }
                    form.DataSources.UserDataSources.Item("SlctAllPn1").Value = "N";
                }
                else
                {
                    for (int i = 0; i < oGrid.Rows.Count; i++)
                    {
                        if (oGrid.DataTable.GetValue("ZERO_IS_INVALID", i).ToString() != "0")
                        {
                            oGrid.DataTable.SetValue("Check", oGrid.GetDataTableRowIndex(i), "N");
                        }
                    }
                    form.DataSources.UserDataSources.Item("SlctAllPn1").Value = "Y";
                }
            }
            catch (Exception e)
            {
                Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private void BackBtn_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                form.PaneLevel--;

                //if (form.PaneLevel == 2)
                //{
                //    form.Items.Item("RefreshBT").Click();
                //}
            }
            catch (Exception e)
            {
                Application.SBO_Application.MessageBox("Message: " + e.Message + "StackTrace: " + e.StackTrace);
            }
        }

        private void CardCodeED_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (string.IsNullOrEmpty(form.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString()))
                {
                    ((SAPbouiCOM.StaticText)form.Items.Item("CardNameST").Specific).Caption = string.Empty;
                }
            }
            catch (Exception e)
            {
                Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace);
            }
        }

        private void CardCodeED_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg chflarg = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string uidChose = chflarg.ChooseFromListUID;
                SAPbouiCOM.DataTable dt = chflarg.SelectedObjects;

                form.DataSources.DBDataSources.Item("OCRD").SetValue("CardCode", 0, chflarg.SelectedObjects.GetValue("CardCode", 0).ToString());
                ((SAPbouiCOM.StaticText)form.Items.Item("CardNameST").Specific).Caption = chflarg.SelectedObjects.GetValue("CardName", 0).ToString();
            }
            catch (Exception e)
            {
                Application.SBO_Application.MessageBox("The Following Error Occrurred:\n" + e.Message + "\n" + e.StackTrace);
            }
        }

        private void NextBtn_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string sSQL,
                   sErr = "A",
                   sCardCode,
                   sElleipsiDS, 
                   sClosedDS; 

            SAPbouiCOM.Grid oFinalGrid;

            try
            {
                Recordset rsError = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsError.DoQuery("INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit', '', '" + company.UserName + "', 'Form: " + form.TypeEx + " " + pVal.FormTypeCount + " item: " + pVal.ItemUID + "', 'SUCCESS')");

                form.Freeze(true);

                if (form.PaneLevel == 1)
                {
                    if (string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("SeriesTo").Value.ToString()))
                    {
                        Application.SBO_Application.StatusBar.SetText("Παρακαλώ Συμπληρώστε Σειρά Παραστατικού Στόχου.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        form.ActiveItem = "SeriesTo";
                        return;
                    }
                    else if (form.DataSources.UserDataSources.Item("SeriesTo").Value.ToString() != "8")
                    {
                        if (Application.SBO_Application.MessageBox("Έχει Επιλεχθεί Σειρά Παραστατικού Στόχου " + ((SAPbouiCOM.ComboBox)form.Items.Item("SeriesTo").Specific).Selected.Description + "\nΣυνέχεια;", 1, "Ναι", "Όχι") != 1)
                        {
                            form.ActiveItem = "SeriesTo";
                            return;
                        }
                    }
                    if (string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("DocDateDS").Value.ToString()))
                    {
                        Application.SBO_Application.StatusBar.SetText("Παρακαλώ Συμπληρώστε Ημ. Καταχώρησης Στόχου.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        form.ActiveItem = "DocDateED";
                        return;
                    }
                }

                //oFinalGrid.CommonSetting.SetRowEditable(0, false);

                sCardCode = form.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString();

                if (form.PaneLevel != 2)
                {
                    form.PaneLevel++;
                }

                if (form.PaneLevel == 2)
                {
                    sErr = "Build sSQL";
                    //                    sSQL = string.Format(Stringia.sNextSQL, form.DataSources.UserDataSources.Item("DocDateDS").Value.ToString(), form.DataSources.UserDataSources.Item("ElleipsiDS").Value.ToString());

                    if (string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("ClosedDS").Value.ToString()))
                    {
                        sClosedDS = "N";
                    }
                    else
                    {
                        sClosedDS = form.DataSources.UserDataSources.Item("ClosedDS").Value.ToString();
                    }

                    if (string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("ElleipsiDS").Value.ToString()))
                    {
                        sElleipsiDS = "N";
                    }
                    else
                    {
                        sElleipsiDS = form.DataSources.UserDataSources.Item("ElleipsiDS").Value.ToString();
                    }

                    sSQL = string.Format(Stringia.sNextSQL, sElleipsiDS, sClosedDS);

                    sSQL += System.Environment.NewLine;

                    string sSeries = "";

                    sErr = "B";
                    for (int i = 0; i < ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).Rows.Count; i++)
                    {
                        if (((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).DataTable.GetValue("Check",
                            ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).GetDataTableRowIndex(i)).ToString() == "Y")
                        {
                            sSeries += "," + ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).DataTable.GetValue("Series",
                                ((SAPbouiCOM.Grid)form.Items.Item("SeriesFrGr").Specific).GetDataTableRowIndex(i)).ToString();
                        }
                    }

                    if (!string.IsNullOrEmpty(sCardCode))
                    {
                        sSQL += " AND Q.\"CardCode\" = '" + sCardCode + "' ";
                    }
                    if (!string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("DocDateFr").Value.ToString()))
                    {
                        sSQL += " AND Q.\"DocDate\" >= '" + form.DataSources.UserDataSources.Item("DocDateFr").Value.ToString() + "' ";
                    }
                    if (!string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("DocDateTo").Value.ToString()))
                    {
                        sSQL += " AND Q.\"DocDate\" <= '" + form.DataSources.UserDataSources.Item("DocDateTo").Value.ToString() + "' ";
                    }
                    if (!string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("TaxDateFr").Value.ToString()))
                    {
                        sSQL += " AND Q.\"TaxDate\" >= '" + form.DataSources.UserDataSources.Item("TaxDateFr").Value.ToString() + "' ";
                    }
                    if (!string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("TaxDateTo").Value.ToString()))
                    {
                        sSQL += " AND Q.\"TaxDate\" <= '" + form.DataSources.UserDataSources.Item("TaxDateTo").Value.ToString() + "' ";
                    }
                    if (sSeries.Length > 0)
                    {
                        sSQL += " AND Q.\"Series\" IN (" + sSeries.Substring(1) + ") ";
                    }
                    if (!string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("SubCatDS").Value.ToString()))
                    {
                        sSQL += " AND IFNULL(SC.\"Name\", '') = '" + form.DataSources.UserDataSources.Item("SubCatDS").Value.ToString() + "' ";
                    }
                    if (!string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("CommentsDS").Value.ToString()))
                    {
                        sSQL += " AND LOWER(IFNULL(TO_VARCHAR(Q.\"U_TKA_BPComments\"), '')) LIKE LOWER('%" + form.DataSources.UserDataSources.Item("CommentsDS").Value.ToString() + "%') ";
                    }
                    if (!string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("ShipTypeDS").Value.ToString()))
                    {
                        sSQL += " AND CASE WHEN IFNULL(Q.\"TrnspCode\", '-1') <> '4' THEN '-1' ELSE IFNULL(Q.\"TrnspCode\", '-1') END= '" + form.DataSources.UserDataSources.Item("ShipTypeDS").Value.ToString() + "' ";
//                        sSQL += " AND Q.\"TrnspCode\" = '" + form.DataSources.UserDataSources.Item("ShipTypeDS").Value.ToString() + "' ";
                    }

                    sSQL += " GROUP BY Q.\"DocNum\", " +
                            "   Q.\"DocEntry\", " +
                            "   Q.\"CardCode\", " +
                            "   Q.\"CardName\", " +
                            "   Q.\"DocDate\", " +
                            "   Q.\"TaxDate\", " +
                            //                            "   '" + form.DataSources.UserDataSources.Item("DocDateDS").Value.ToString() + "', " +
                            "   Q.\"DocTotal\", " +
                            "   C.\"LicTradNum\", " +
                            "   C.\"U_TKA_CustSubCategory\", " +
                            "   TO_VARCHAR(Q.\"U_TKA_BPComments\"), " +
                            "   N1.\"SeriesName\", " +
                            "   IFNULL(A.A, 0), " +
                            "   CASE IFNULL(Q.\"U_TKA_WebOrder\", '') WHEN '' THEN Q.\"NumAtCard\" ELSE Q.\"U_TKA_WebOrder\" END " +
                            " ORDER BY 2";

                    oFinalGrid = ((SAPbouiCOM.Grid)form.Items.Item("ResultsGRD").Specific);

                    //System.IO.File.WriteAllText(@"C:\Users\wm.user1\Desktop\sSQL.txt", sSQL);
                    
                    oFinalGrid.DataTable.ExecuteQuery(sSQL);

                    sErr = "Set oFinalGrid Columns";
                    for (int i = 0; i < oFinalGrid.DataTable.Columns.Count; i++)
                    {
                        oFinalGrid.Columns.Item(i).Editable = false;
                        oFinalGrid.Columns.Item(i).TitleObject.Sortable = true;
                    }

                    oFinalGrid.Columns.Item("Check").Editable = true;

                    oFinalGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

                    oFinalGrid.Columns.Item("DocTotal").RightJustified = true;

                    ((SAPbouiCOM.EditTextColumn)oFinalGrid.Columns.Item("CardCode")).LinkedObjectType = "2";
                    ((SAPbouiCOM.EditTextColumn)oFinalGrid.Columns.Item("DocEntry")).LinkedObjectType = "23";

                    oFinalGrid.Columns.Item("DocEntry").TitleObject.Caption = "Κλειδί Παρ/κού";
                    oFinalGrid.Columns.Item("DocNum").TitleObject.Caption = "Αριθμός";
                    oFinalGrid.Columns.Item("CardCode").TitleObject.Caption = "Κωδικός Πελάτη";
                    oFinalGrid.Columns.Item("CardName").TitleObject.Caption = "Επωνυμία Πελάτη";
                    oFinalGrid.Columns.Item("LicTradNum").TitleObject.Caption = "ΑΦΜ Πελάτη";
                    oFinalGrid.Columns.Item("U_TKA_CustSubCategory").TitleObject.Caption = "Οικογένεια Πελάτη";
                    oFinalGrid.Columns.Item("DocTotal").TitleObject.Caption = "Συνολική Αξία";
                    oFinalGrid.Columns.Item("NumAtCard").TitleObject.Caption = "Αριθμός Αναφοράς";
                    oFinalGrid.Columns.Item("DocDate").TitleObject.Caption = "Ημ. Έκδοσης";
                    oFinalGrid.Columns.Item("TaxDate").TitleObject.Caption = "Ημ. Καταχώρησης";
                    oFinalGrid.Columns.Item("SeriesName").TitleObject.Caption = "Σειρά παραστατικού";
                    oFinalGrid.Columns.Item("U_TKA_BPComments").TitleObject.Caption = "Σχόλια Πελατών";

                    //oFinalGrid.Columns.Item("QryGroup4").Visible = false;
                    oFinalGrid.Columns.Item("DocDate").Visible = false;
                    oFinalGrid.Columns.Item("ZERO_IS_INVALID").Visible = false;

                    oFinalGrid.Columns.Item("DocNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);

                    for (int i = 0; i < oFinalGrid.DataTable.Rows.Count; i++)
                    {
                        if (oFinalGrid.DataTable.GetValue("ZERO_IS_INVALID", i).ToString() == "0"
                            || form.DataSources.UserDataSources.Item("ShipTypeDS").Value.ToString() == "4")
                        {
                            oFinalGrid.CommonSetting.SetRowEditable(i + 1, false);
                        }
                    }

                    oFinalGrid.AutoResizeColumns();
                }

                form.DataSources.UserDataSources.Item("SlctAllPn1").Value = "Y";
            }
            catch (Exception e)
            {
                Application.SBO_Application.MessageBox("Παρουσιάστηκε το Παρακάτω Σφάλμα:\n" + e.Message + "\n" + e.StackTrace + "\nsErr: " + sErr);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private static void ExecBtn_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string sErr = "";

            try
            {
                Recordset rsError = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsError.DoQuery("INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit', '', '" + company.UserName + "', 'Form: " + form.TypeEx + " " + pVal.FormTypeCount + " item: " + pVal.ItemUID + "', 'SUCCESS')");

                form.Freeze(true);

                string sDocEntries = "0",
                   sError,
                   sSQL,
                   sCardCode,
                   sItemType,
                   sNewDocEntry;

                SAPbouiCOM.Grid oGrid,
                                oAddResultGrid;

                SAPbouiCOM.DataTable oResultDT;

                bool bSelected = false,
                     bDontAdd = false,
                     bElleipiEidi = false;

                Recordset rsGetData,
                          rsErrorLog,
                          rsGetPFSData,
                          rsCheckListNum4,
                          rsInsertIntoErrorsTable;

                Documents oOrder,
                          oOrderPFS,
                          oBaseDoc = (Documents)company.GetBusinessObject(BoObjectTypes.oQuotations),
                          oBaseDocPFS;

                int iDocEntry,
                    iProperty;

                DateTime dtDocDate;

                BusinessPartners oBP = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (Application.SBO_Application.MessageBox("Δημιουργία Παραγγελιών για τις επιλεγμένες γραμμές;", 2, "Ναι", "Όχι") != 1)
                {
                    return;
                }
                rsInsertIntoErrorsTable = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                //                rsInsertIntoErrorsTable.DoQuery("INSERT INTO ERRORS SELECT current_date || ' ' || current_time, 'SalesOrderSplit', '', '', 'Tran Start', 'SUCCESS' FROM DUMMY;");

                sErr = "A";
                oGrid = (SAPbouiCOM.Grid)form.Items.Item("ResultsGRD").Specific;

                sErr = "B";
                for (int i = 0; i < oGrid.Rows.Count; i++)
                {
                    if (oGrid.DataTable.GetValue("Check", oGrid.GetDataTableRowIndex(i)).ToString() == "Y")
                    {
                        sDocEntries += "," + oGrid.DataTable.GetValue("DocEntry", oGrid.GetDataTableRowIndex(i)).ToString();

                        bSelected = true;
                    }
                }

                sErr = "C";
                if (!bSelected)
                {
                    Application.SBO_Application.MessageBox("Επιλέξτε τουλάχιστον μία γραμμη.");
                    return;
                }

                sErr = "D";
                oAddResultGrid = (SAPbouiCOM.Grid)form.Items.Item("AddResGRD").Specific;

                oResultDT = oAddResultGrid.DataTable;

                oResultDT.Rows.Clear();

                sSQL = "SELECT COUNT(*) A FROM QUT1 Q1 " +
                       " INNER JOIN OITM I ON I.\"ItemCode\" = Q1.\"ItemCode\" " +
                       " WHERE Q1.\"DocEntry\" IN (" + sDocEntries + ") " +
                       " AND I.\"QryGroup4\" = 'Y'";

                rsGetData = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsGetData.DoQuery(sSQL);

                if (Convert.ToInt32(rsGetData.Fields.Item(0).Value) != 0)
                {
                    if (Application.SBO_Application.MessageBox("Βρέθηκαν είδη σε έλλειψη.\nΣυνέχεια;", 2, "Ναι", "Όχι") != 1)
                    {
                        return;
                    }

                    bElleipiEidi = true;
                }

                sSQL = string.Format(Stringia.sExecSQL, sDocEntries);

                //System.IO.File.WriteAllText(@"C:\Users\wm.user1\Desktop\sSQL.txt", sSQL);

                rsGetData.DoQuery(sSQL);

                if (rsGetData.RecordCount <= 0)
                {
                    Application.SBO_Application.MessageBox("Δεδομένα δεν βρέθηκαν με τα συγκεκριμένα κριτήρια.");
                    return;
                }

                SAPbouiCOM.ProgressBar prg = Application.SBO_Application.StatusBar.CreateProgressBar("sdasda", rsGetData.RecordCount - 1, false);
                prg.Text = "Καταχώρηση Παραστατικών";
                prg.Value = 0;

                rsGetData.MoveFirst();

                iProperty = Convert.ToInt32(rsGetData.Fields.Item("PROPERTIES").Value);
                sCardCode = rsGetData.Fields.Item("CardCode").Value.ToString();
                iDocEntry = Convert.ToInt32(rsGetData.Fields.Item("DocEntry").Value);
                sItemType = rsGetData.Fields.Item("ItmsGrpNam").Value.ToString();

                if (iDocEntry != 99999999)
                {
                    oBaseDoc.GetByKey(iDocEntry);
                }
                else
                {
                    return;
                }

                sErr = "Get Correct DateTime formats";
                SBObob oSBObob = (SBObob)company.GetBusinessObject(BoObjectTypes.BoBridge);
                Recordset rsDateTimeFormat = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsDateTimeFormat = oSBObob.Format_StringToDate(form.DataSources.UserDataSources.Item("DocDateDS").ValueEx);
                dtDocDate = Convert.ToDateTime(rsDateTimeFormat.Fields.Item(0).Value);

                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDateTimeFormat);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);

                    rsDateTimeFormat = null;
                    oSBObob = null;
                }
                catch (Exception) { }

                oOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);

                while (!rsGetData.EoF)
                {
                    if (iProperty != Convert.ToInt32(rsGetData.Fields.Item("PROPERTIES").Value)
                        || sCardCode != rsGetData.Fields.Item("CardCode").Value.ToString()
                        || iDocEntry != Convert.ToInt32(rsGetData.Fields.Item("DocEntry").Value))
                    {
                        iProperty = Convert.ToInt32(rsGetData.Fields.Item("PROPERTIES").Value);
                        sCardCode = rsGetData.Fields.Item("CardCode").Value.ToString();
                        iDocEntry = Convert.ToInt32(rsGetData.Fields.Item("DocEntry").Value);

                        oBP.GetByKey(oOrder.CardCode);

                        if (oBP.PriceListNum == 4)
                        {
                            rsCheckListNum4 = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            rsCheckListNum4.DoQuery("SELECT COUNT(*) A FROM QUT1 Q1 " +
                                                    " INNER JOIN OITM I ON I.\"ItemCode\" = Q1.\"ItemCode\" " +
                                                    " WHERE I.\"QryGroup4\" = 'Y' " +
                                                    " AND Q1.\"DocEntry\" = " + oBaseDoc.DocEntry);

                            if (rsCheckListNum4.Fields.Item("A").Value.ToString() != "0")
                            {
                                bDontAdd = true;
                            }

                            try
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsCheckListNum4);

                                rsCheckListNum4 = null;
                            }
                            catch (Exception) { }
                        }


                        if (bDontAdd)
                        {
                            oResultDT.Rows.Add();

                            oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Σφάλμα");
                            oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, oBaseDoc.DocEntry);
                            oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                            oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, -1);
                            oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, "17");
                            oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, "-23");
                            oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, "Σε τιμοκατάλογο Φαρμακεία δεν επιτρέπεται διαχωρισμός με είδη σε έλλειψη.");
                            oResultDT.SetValue("ItmType", oResultDT.Rows.Count - 1, sItemType);
                        }
                        else
                        {
                            if (oOrder.Add() != 0)
                            {
                                oResultDT.Rows.Add();

                                oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Σφάλμα");
                                oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, oBaseDoc.DocEntry);
                                oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                                oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, -1);
                                oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, "17");
                                oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, company.GetLastErrorCode());
                                oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, company.GetLastErrorDescription());
                                oResultDT.SetValue("ItmType", oResultDT.Rows.Count - 1, sItemType);
                            }
                            else
                            {
                                sNewDocEntry = company.GetNewObjectKey();

                                oResultDT.Rows.Add();

                                oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Επιτυχία");
                                oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, oBaseDoc.DocEntry);
                                oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                                oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, sNewDocEntry);
                                oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, "17");
                                oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, -1);
                                oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, string.Empty);
                                oResultDT.SetValue("ItmType", oResultDT.Rows.Count - 1, sItemType);

                                sErr = "Update Order UDFS";
                                rsGetPFSData = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rsGetPFSData.DoQuery("CALL TKA_SP_UPDATE_ORDR_UDFS_FROM_OQUT(" + sNewDocEntry + ")");

                                sErr = "Check for Existing PFS on quotation";
                                rsGetPFSData.DoQuery("SELECT \"VisOrder\" " +
                                                     "   FROM QUT1 " +
                                                     "   WHERE LEFT(\"ItemCode\", 3) = 'ΠΦΣ' " +
                                                     "   AND \"DocEntry\" = " + oBaseDoc.DocEntry);

                                if (rsGetPFSData.RecordCount > 0)
                                {
                                    oBaseDocPFS = (Documents)company.GetBusinessObject(BoObjectTypes.oQuotations);
                                    oBaseDocPFS.GetByKey(oBaseDoc.DocEntry);

                                    while (!rsGetPFSData.EoF)
                                    {
                                        oBaseDocPFS.Lines.SetCurrentLine(Convert.ToInt32(rsGetPFSData.Fields.Item(0).Value));

                                        oBaseDocPFS.Lines.LineStatus = BoStatus.bost_Close;

                                        rsGetPFSData.MoveNext();
                                    }

                                    if (oBaseDocPFS.Update() != 0)
                                    {
                                        rsErrorLog = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        sError = "INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit on Update Base Doc for PFS', '', '" + company.UserName + "', '" + company.GetLastErrorDescription().Replace("'", "") + " DocEntry: " + oBaseDocPFS.DocEntry + "', 'ERROR')";
                                        rsErrorLog.DoQuery(sError);
                                    }
                                }

                                //check for PFS on Order
                                rsGetPFSData.DoQuery("SELECT ERROR_ID, ITEM_CODE, AMOUNT, VAT_GROUP FROM TKA_F_CALCULATE_PFS('17', " + sNewDocEntry + ");");

                                if (rsGetPFSData.RecordCount > 0)
                                {
                                    rsGetPFSData.MoveFirst();

                                    oOrderPFS = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);
                                    oOrderPFS.GetByKey(Convert.ToInt32(sNewDocEntry));

                                    while (!rsGetPFSData.EoF)
                                    {
                                        if (Convert.ToInt32(rsGetPFSData.Fields.Item("ERROR_ID").Value, CultureInfo.InvariantCulture) != 0)
                                        {
                                            oOrderPFS.Lines.Add();

                                            oOrderPFS.Lines.ItemCode = Convert.ToString(rsGetPFSData.Fields.Item("ITEM_CODE").Value, CultureInfo.InvariantCulture);
                                            oOrderPFS.Lines.LineTotal = Convert.ToDouble(rsGetPFSData.Fields.Item("AMOUNT").Value, CultureInfo.InvariantCulture);
                                            oOrderPFS.Lines.VatGroup = Convert.ToString(rsGetPFSData.Fields.Item("VAT_GROUP").Value, CultureInfo.InvariantCulture);
                                        }

                                        rsGetPFSData.MoveNext();
                                    }

                                    if (oOrderPFS.Update() != 0)
                                    {
                                        rsErrorLog = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        sError = "INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit on Update for PFS', '', '" + company.UserName + "', '" + company.GetLastErrorDescription().Replace("'", "") + " DocEntry: " + oOrderPFS.DocEntry + "', 'ERROR')";
                                        //Application.MessageBox("[2] Failed to update document...\n\nError Code: " + company.GetLastErrorCode() + "\nError Message: " + company.GetLastErrorDescription());
                                        rsErrorLog.DoQuery(sError);
                                    }
                                }

                                try
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsGetPFSData);

                                    rsGetPFSData = null;
                                }
                                catch (Exception) { }
                            }
                        }

                        bDontAdd = false;

                        sItemType = rsGetData.Fields.Item("ItmsGrpNam").Value.ToString();

                        oOrder = null;
                        oOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);

                        if (iDocEntry != 99999999)
                        {
                            oBaseDoc.GetByKey(iDocEntry);
                        }
                        else
                        {
                        }
                    }

                    if (iDocEntry == 99999999)
                    {
                        break;
                    }

                    oOrder.Series = Convert.ToInt32(form.DataSources.UserDataSources.Item("SeriesTo").Value);
                    oOrder.CardCode = oBaseDoc.CardCode;
                    oOrder.NumAtCard = oBaseDoc.NumAtCard;
                    oOrder.DocDate = dtDocDate;
                    oOrder.TaxDate = dtDocDate;
                    oOrder.DocDueDate = dtDocDate;

                    oBaseDoc.Lines.SetCurrentLine(Convert.ToInt32(rsGetData.Fields.Item("VisOrder").Value));

                    oOrder.Lines.BaseLine = Convert.ToInt32(rsGetData.Fields.Item("LineNum").Value);
                    oOrder.Lines.BaseEntry = iDocEntry;
                    oOrder.Lines.BaseType = 23;

                    oOrder.Lines.Add();

                    prg.Value++;

                    rsGetData.MoveNext();
                }

                try
                {
                    if (prg != null)
                    {
                        prg.Stop();
                    }
                }
                catch (Exception)
                { }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(prg);
                    prg = null;
                }

                // edw fernoume apo katw ta elleiph eidh apo ta epilegmena quotations, ean exei epilexthei to antistoixo checkbox
                if (bElleipiEidi)
                {
                    try
                    {
                        /*sSQL = "SELECT DISTINCT Q1.\"DocEntry\" FROM QUT1 Q1 " +
                               " INNER JOIN OITM I ON I.\"ItemCode\" = Q1.\"ItemCode\" " +
                               " WHERE Q1.\"DocEntry\" IN (" + sDocEntries + ") " +
                               " AND I.\"QryGroup4\" = 'Y' ORDER BY 1 ASC";*/

                        sSQL = "SELECT Q1.\"DocEntry\", STRING_AGG(I.\"ItemCode\", ', ') ITMS FROM QUT1 Q1 " +
                               "    INNER JOIN OITM I ON I.\"ItemCode\" = Q1.\"ItemCode\" " +
                               "    WHERE 1 = 1 " +
                               "    AND Q1.\"DocEntry\" IN(" + sDocEntries + ") " +
                               "    AND I.\"QryGroup4\" = 'Y' " +
                               "    GROUP BY  Q1.\"DocEntry\" ORDER BY 1 ASC ";

                        rsGetData = null;
                        rsGetData = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsGetData.DoQuery(sSQL);

                        while (!rsGetData.EoF)
                        {
                            oResultDT.Rows.Add();

                            oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Ελλειπή Είδη");
                            oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, rsGetData.Fields.Item("DocEntry").Value.ToString());
                            oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                            oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, -1);
                            oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, string.Empty);
                            oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, -1);
                            oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, rsGetData.Fields.Item("ITMS").Value.ToString());
                            oResultDT.SetValue("ItmType", oResultDT.Rows.Count - 1, "Είδη σε Έλλειψη");

                            rsGetData.MoveNext();
                        }
                    }
                    catch (Exception)
                    {
                        //Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace);
                    }
                }

                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsGetData);

                    rsGetData = null;
                }
                catch (Exception) { }

                oAddResultGrid.Columns.Item("AddRes").TitleObject.Caption = "Αποτέλεσμα";
                oAddResultGrid.Columns.Item("OrgDocNtr").TitleObject.Caption = "Κλ.Παραστατικού Βάσης";
                oAddResultGrid.Columns.Item("OriginType").TitleObject.Caption = "Τύπος Παραστατικού Βάσης";
                oAddResultGrid.Columns.Item("TrgDocNtry").TitleObject.Caption = "Κλ.Παραστατικού Στόχου";
                oAddResultGrid.Columns.Item("TargetType").TitleObject.Caption = "Τύπος Παραστατικού Στόχου";
                oAddResultGrid.Columns.Item("SapErrCode").TitleObject.Caption = "Κωδ.Σφάλματος";
                oAddResultGrid.Columns.Item("SapErrMsg").TitleObject.Caption = "Περιγραφή Σφάλματος";
                oAddResultGrid.Columns.Item("ItmType").TitleObject.Caption = "Τύπος Είδους";

                oAddResultGrid.Columns.Item("OrgObjTp").Visible = false;
                oAddResultGrid.Columns.Item("OriginType").Visible = false;
                oAddResultGrid.Columns.Item("TargetType").Visible = false;

                ((SAPbouiCOM.EditTextColumn)oAddResultGrid.Columns.Item("OrgDocNtr")).LinkedObjectType = "23";
                ((SAPbouiCOM.EditTextColumn)oAddResultGrid.Columns.Item("TrgDocNtry")).LinkedObjectType = "17";

                form.PaneLevel++;

                oAddResultGrid.AutoResizeColumns();

                //                rsInsertIntoErrorsTable.DoQuery("INSERT INTO ERRORS SELECT current_date || ' ' || current_time, 'SalesOrderSplit', '', '', 'Tran End', 'SUCCESS' FROM DUMMY;");
            }
            catch (Exception e)
            {
                Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace + "\nAt: " + sErr);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private static void New_ExecBtn_ClickAfter(object sboObject, SAPbouiCOM.ItemEvent pVal)
        {
            string sErr = "";

            try
            {
                Recordset rsError = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsError.DoQuery("INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit', '', '" + company.UserName + "', 'Form: " + form.TypeEx + " " + pVal.FormTypeCount + " item: " + pVal.ItemUID + "', 'SUCCESS')");

                form.Freeze(true);

                string sDocEntries = "0",
                   sError,
                   sSQL,
                   sCardCode,
                   sItemType,
                   sNewDocEntry;

                SAPbouiCOM.Grid oGrid,
                                oAddResultGrid;

                SAPbouiCOM.DataTable oResultDT;

                bool bSelected = false,
                     bDontAdd = false,
                     bElleipiEidi = false;

                Recordset rsGetData,
                          rsErrorLog,
                          rsGetPFSData,
                          rsCheckListNum4,
                          rsInsertIntoErrorsTable;

                Documents oOrder,
                          oOrderPFS,
                          oBaseDoc = (Documents)company.GetBusinessObject(BoObjectTypes.oQuotations),
                          oBaseDocPFS;

                int iDocEntry,
                    iProperty;

                DateTime dtDocDate;

                BusinessPartners oBP = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (Application.SBO_Application.MessageBox("Δημιουργία Παραγγελιών για τις επιλεγμένες γραμμές;", 2, "Ναι", "Όχι") != 1)
                {
                    return;
                }
                rsInsertIntoErrorsTable = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                //                rsInsertIntoErrorsTable.DoQuery("INSERT INTO ERRORS SELECT current_date || ' ' || current_time, 'SalesOrderSplit', '', '', 'Tran Start', 'SUCCESS' FROM DUMMY;");

                sErr = "A";
                oGrid = (SAPbouiCOM.Grid)form.Items.Item("ResultsGRD").Specific;

                sErr = "B";
                for (int i = 0; i < oGrid.Rows.Count; i++)
                {
                    if (oGrid.DataTable.GetValue("Check", oGrid.GetDataTableRowIndex(i)).ToString() == "Y")
                    {
                        sDocEntries += "," + oGrid.DataTable.GetValue("DocEntry", oGrid.GetDataTableRowIndex(i)).ToString();

                        bSelected = true;
                    }
                }

                sErr = "C";
                if (!bSelected)
                {
                    Application.SBO_Application.MessageBox("Επιλέξτε τουλάχιστον μία γραμμη.");
                    return;
                }

                sErr = "D";
                oAddResultGrid = (SAPbouiCOM.Grid)form.Items.Item("AddResGRD").Specific;

                oResultDT = oAddResultGrid.DataTable;

                oResultDT.Rows.Clear();

                sSQL = "SELECT COUNT(*) A FROM QUT1 Q1 " +
                       " INNER JOIN OITM I ON I.\"ItemCode\" = Q1.\"ItemCode\" " +
                       " WHERE Q1.\"DocEntry\" IN (" + sDocEntries + ") " +
                       " AND I.\"QryGroup4\" = 'Y'";

                rsGetData = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsGetData.DoQuery(sSQL);

                if (Convert.ToInt32(rsGetData.Fields.Item(0).Value) != 0)
                {
                    if (Application.SBO_Application.MessageBox("Βρέθηκαν είδη σε έλλειψη.\nΣυνέχεια;", 2, "Ναι", "Όχι") != 1)
                    {
                        return;
                    }

                    bElleipiEidi = true;
                }

                sSQL = string.Format(Stringia.sExecSQL, sDocEntries);

                //System.IO.File.WriteAllText(@"C:\Users\wm.user1\Desktop\sSQL.txt", sSQL);

                rsGetData.DoQuery(sSQL);

                if (rsGetData.RecordCount <= 0)
                {
                    Application.SBO_Application.MessageBox("Δεδομένα δεν βρέθηκαν με τα συγκεκριμένα κριτήρια.");
                    form.Freeze(false);
                    return;
                }

                SAPbouiCOM.ProgressBar prg = Application.SBO_Application.StatusBar.CreateProgressBar("sdasda", rsGetData.RecordCount - 1, false);
                prg.Text = "Καταχώρηση Παραστατικών";
                prg.Value = 0;

                rsGetData.MoveFirst();

                iProperty = Convert.ToInt32(rsGetData.Fields.Item("PROPERTIES").Value);
                sCardCode = rsGetData.Fields.Item("CardCode").Value.ToString();
                iDocEntry = Convert.ToInt32(rsGetData.Fields.Item("DocEntry").Value);
                sItemType = rsGetData.Fields.Item("ItmsGrpNam").Value.ToString();

                if (iDocEntry != 99999999)
                {
                    oBaseDoc.GetByKey(iDocEntry);
                }
                else
                {
                    Application.SBO_Application.MessageBox("Δεδομένα δεν βρέθηκαν με τα συγκεκριμένα κριτήρια.");
                    form.Freeze(false);
                    return;
                }

                sErr = "Get Correct DateTime formats";
                SBObob oSBObob = (SBObob)company.GetBusinessObject(BoObjectTypes.BoBridge);
                Recordset rsDateTimeFormat = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsDateTimeFormat = oSBObob.Format_StringToDate(form.DataSources.UserDataSources.Item("DocDateDS").ValueEx);
                dtDocDate = Convert.ToDateTime(rsDateTimeFormat.Fields.Item(0).Value);

                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDateTimeFormat);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSBObob);

                    rsDateTimeFormat = null;
                    oSBObob = null;
                }
                catch (Exception) { }

                oOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);

                while (!rsGetData.EoF)
                {
                    if (iProperty != Convert.ToInt32(rsGetData.Fields.Item("PROPERTIES").Value)
                        || sCardCode != rsGetData.Fields.Item("CardCode").Value.ToString()
                        || iDocEntry != Convert.ToInt32(rsGetData.Fields.Item("DocEntry").Value))
                    {
                        iProperty = Convert.ToInt32(rsGetData.Fields.Item("PROPERTIES").Value);
                        sCardCode = rsGetData.Fields.Item("CardCode").Value.ToString();
                        iDocEntry = Convert.ToInt32(rsGetData.Fields.Item("DocEntry").Value);

                        oBP.GetByKey(oOrder.CardCode);

                        if (oBP.PriceListNum == 4)
                        {
                            rsCheckListNum4 = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            rsCheckListNum4.DoQuery("SELECT COUNT(*) A FROM QUT1 Q1 " +
                                                    " INNER JOIN OITM I ON I.\"ItemCode\" = Q1.\"ItemCode\" " +
                                                    " WHERE I.\"QryGroup4\" = 'Y' " +
                                                    " AND Q1.\"DocEntry\" = " + oBaseDoc.DocEntry);

                            if (rsCheckListNum4.Fields.Item("A").Value.ToString() != "0")
                            {
                                bDontAdd = true;
                            }

                            try
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsCheckListNum4);

                                rsCheckListNum4 = null;
                            }
                            catch (Exception) { }
                        }


                        if (bDontAdd)
                        {
                            oResultDT.Rows.Add();

                            oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Σφάλμα");
                            oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, oBaseDoc.DocEntry);
                            oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                            oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, -1);
                            oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, "17");
                            oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, "-23");
                            oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, "Σε τιμοκατάλογο Φαρμακεία δεν επιτρέπεται διαχωρισμός με είδη σε έλλειψη.");
                            oResultDT.SetValue("ItmType", oResultDT.Rows.Count - 1, sItemType);
                        }
                        else
                        {
                            if (oOrder.Add() != 0)
                            {
                                oResultDT.Rows.Add();

                                oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Σφάλμα");
                                oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, oBaseDoc.DocEntry);
                                oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                                oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, -1);
                                oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, "17");
                                oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, company.GetLastErrorCode());
                                oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, company.GetLastErrorDescription());
                                oResultDT.SetValue("ItmType", oResultDT.Rows.Count - 1, sItemType);
                            }
                            else
                            {
                                sNewDocEntry = company.GetNewObjectKey();

                                /* DEBUG */
                                rsErrorLog = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                sError = "INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplitDEBUG', '', '', 'GetNewObjectKey: " + sNewDocEntry + "', 'DEBUG')";
                                rsErrorLog.DoQuery(sError);
                                /* DEBUG */

                                oResultDT.Rows.Add();

                                oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Επιτυχία");
                                oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, oBaseDoc.DocEntry);
                                oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                                oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, sNewDocEntry);
                                oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, "17");
                                oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, -1);
                                oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, string.Empty);
                                oResultDT.SetValue("ItmType", oResultDT.Rows.Count - 1, sItemType);

                                sErr = "Update Order UDFS";
                                rsGetPFSData = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rsGetPFSData.DoQuery("CALL TKA_SP_UPDATE_ORDR_UDFS_FROM_OQUT(" + sNewDocEntry + ")");

                                sErr = "Check for Existing PFS on quotation";
                                rsGetPFSData = null;
                                rsGetPFSData = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rsGetPFSData.DoQuery("SELECT \"VisOrder\" " +
                                                     "   FROM QUT1 " +
                                                     "   WHERE LEFT(\"ItemCode\", 3) = 'ΠΦΣ' " +
                                                     "   AND \"DocEntry\" = " + oBaseDoc.DocEntry);

                                /* DEBUG */
                                rsErrorLog = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                sError = "INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplitDEBUG', '', '', 'BaseDocPFSCount: " + rsGetPFSData.RecordCount + " oBaseDoc.DocEntry: " + oBaseDoc.DocEntry + "', 'DEBUG')";
                                rsErrorLog.DoQuery(sError);
                                /* DEBUG */

                                if (rsGetPFSData.RecordCount > 0)
                                {
                                    oBaseDocPFS = (Documents)company.GetBusinessObject(BoObjectTypes.oQuotations);
                                    oBaseDocPFS.GetByKey(oBaseDoc.DocEntry);

                                    while (!rsGetPFSData.EoF)
                                    {
                                        oBaseDocPFS.Lines.SetCurrentLine(Convert.ToInt32(rsGetPFSData.Fields.Item(0).Value));

                                        oBaseDocPFS.Lines.LineStatus = BoStatus.bost_Close;

                                        rsGetPFSData.MoveNext();
                                    }

                                    if (oBaseDocPFS.Update() != 0)
                                    {
                                        rsErrorLog = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        sError = "INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit on Update Base Doc for PFS', '', '" + company.UserName + "', '" + company.GetLastErrorDescription().Replace("'", "") + " DocEntry: " + oBaseDocPFS.DocEntry + "', 'ERROR')";
                                        rsErrorLog.DoQuery(sError);
                                    }
                                }

                                sErr = "check for PFS on Order";
                                rsGetPFSData = null;
                                rsGetPFSData = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                rsGetPFSData.DoQuery("SELECT ERROR_ID, ITEM_CODE, AMOUNT, VAT_GROUP FROM TKA_F_CALCULATE_PFS('17', " + sNewDocEntry + ");");

                                /* DEBUG */
                                rsErrorLog = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                sError = "INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplitDEBUG', '', '', 'NewDocPFSCount: " + rsGetPFSData.RecordCount + " sNewDocEntry: " + sNewDocEntry + "', 'DEBUG')";
                                rsErrorLog.DoQuery(sError);
                                /* DEBUG */

                                if (rsGetPFSData.RecordCount > 0)
                                {
                                    rsGetPFSData.MoveFirst();

                                    oOrderPFS = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);
                                    oOrderPFS.GetByKey(Convert.ToInt32(sNewDocEntry));

                                    while (!rsGetPFSData.EoF)
                                    {
                                        if (Convert.ToInt32(rsGetPFSData.Fields.Item("ERROR_ID").Value, CultureInfo.InvariantCulture) != 0)
                                        {
                                            oOrderPFS.Lines.Add();

                                            oOrderPFS.Lines.ItemCode = Convert.ToString(rsGetPFSData.Fields.Item("ITEM_CODE").Value, CultureInfo.InvariantCulture);
                                            oOrderPFS.Lines.LineTotal = Convert.ToDouble(rsGetPFSData.Fields.Item("AMOUNT").Value, CultureInfo.InvariantCulture);
                                            oOrderPFS.Lines.VatGroup = Convert.ToString(rsGetPFSData.Fields.Item("VAT_GROUP").Value, CultureInfo.InvariantCulture);
                                        }

                                        rsGetPFSData.MoveNext();
                                    }

                                    if (oOrderPFS.Update() != 0)
                                    {
                                        rsErrorLog = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        sError = "INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit on Update for PFS', '', '" + company.UserName + "', '" + company.GetLastErrorDescription().Replace("'", "") + " DocEntry: " + oOrderPFS.DocEntry + "', 'ERROR')";
                                        //Application.MessageBox("[2] Failed to update document...\n\nError Code: " + company.GetLastErrorCode() + "\nError Message: " + company.GetLastErrorDescription());
                                        rsErrorLog.DoQuery(sError);
                                    }
                                }

                                try
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsGetPFSData);

                                    rsGetPFSData = null;
                                }
                                catch (Exception) { }
                            }
                        }

                        bDontAdd = false;

                        sItemType = rsGetData.Fields.Item("ItmsGrpNam").Value.ToString();

                        oOrder = null;
                        oOrder = (Documents)company.GetBusinessObject(BoObjectTypes.oOrders);

                        if (iDocEntry != 99999999)
                        {
                            oBaseDoc.GetByKey(iDocEntry);
                        }
                        else
                        {
                        }
                    }

                    if (iDocEntry == 99999999)
                    {
                        break;
                    }

                    oOrder.Series = Convert.ToInt32(form.DataSources.UserDataSources.Item("SeriesTo").Value);
                    oOrder.CardCode = oBaseDoc.CardCode;
                    oOrder.NumAtCard = oBaseDoc.NumAtCard;
                    oOrder.DocDate = dtDocDate;
                    oOrder.TaxDate = dtDocDate;
                    oOrder.DocDueDate = dtDocDate;

                    oBaseDoc.Lines.SetCurrentLine(Convert.ToInt32(rsGetData.Fields.Item("VisOrder").Value));

                    oOrder.Lines.BaseLine = Convert.ToInt32(rsGetData.Fields.Item("LineNum").Value);
                    oOrder.Lines.BaseEntry = iDocEntry;
                    oOrder.Lines.BaseType = 23;

                    oOrder.Lines.Add();

                    prg.Value++;

                    rsGetData.MoveNext();
                }

                try
                {
                    if (prg != null)
                    {
                        prg.Stop();
                    }
                }
                catch (Exception)
                { }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(prg);
                    prg = null;
                }

                // edw fernoume apo katw ta elleiph eidh apo ta epilegmena quotations, ean exei epilexthei to antistoixo checkbox
                if (bElleipiEidi)
                {
                    try
                    {
                        /*sSQL = "SELECT DISTINCT Q1.\"DocEntry\" FROM QUT1 Q1 " +
                               " INNER JOIN OITM I ON I.\"ItemCode\" = Q1.\"ItemCode\" " +
                               " WHERE Q1.\"DocEntry\" IN (" + sDocEntries + ") " +
                               " AND I.\"QryGroup4\" = 'Y' ORDER BY 1 ASC";*/


                        sSQL = "SELECT Q1.\"DocEntry\", STRING_AGG(I.\"ItemCode\", ', ') ITMS FROM QUT1 Q1 " +
                               "    INNER JOIN OITM I ON I.\"ItemCode\" = Q1.\"ItemCode\" " +
                               "    WHERE 1 = 1 " +
                               "    AND Q1.\"DocEntry\" IN(" + sDocEntries + ") " +
                               "    AND I.\"QryGroup4\" = 'Y' " +
                               "    GROUP BY  Q1.\"DocEntry\" ORDER BY 1 ASC ";

                        rsGetData = null;
                        rsGetData = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsGetData.DoQuery(sSQL);

                        while (!rsGetData.EoF)
                        {
                            oResultDT.Rows.Add();

                            oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Ελλειπή Είδη");
                            oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, rsGetData.Fields.Item("DocEntry").Value.ToString());
                            oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                            oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, -1);
                            oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, string.Empty);
                            oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, -1);
                            oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, rsGetData.Fields.Item("ITMS").Value.ToString());
                            oResultDT.SetValue("ItmType", oResultDT.Rows.Count - 1, "Είδη σε Έλλειψη");

                            rsGetData.MoveNext();
                        }
                    }
                    catch (Exception)
                    {
                        //Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace);
                    }
                }

                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsGetData);

                    rsGetData = null;
                }
                catch (Exception) { }

                oAddResultGrid.Columns.Item("AddRes").TitleObject.Caption = "Αποτέλεσμα";
                oAddResultGrid.Columns.Item("OrgDocNtr").TitleObject.Caption = "Κλ.Παραστατικού Βάσης";
                oAddResultGrid.Columns.Item("OriginType").TitleObject.Caption = "Τύπος Παραστατικού Βάσης";
                oAddResultGrid.Columns.Item("TrgDocNtry").TitleObject.Caption = "Κλ.Παραστατικού Στόχου";
                oAddResultGrid.Columns.Item("TargetType").TitleObject.Caption = "Τύπος Παραστατικού Στόχου";
                oAddResultGrid.Columns.Item("SapErrCode").TitleObject.Caption = "Κωδ.Σφάλματος";
                oAddResultGrid.Columns.Item("SapErrMsg").TitleObject.Caption = "Περιγραφή Σφάλματος";
                oAddResultGrid.Columns.Item("ItmType").TitleObject.Caption = "Τύπος Είδους";

                oAddResultGrid.Columns.Item("OrgObjTp").Visible = false;
                oAddResultGrid.Columns.Item("OriginType").Visible = false;
                oAddResultGrid.Columns.Item("TargetType").Visible = false;

                ((SAPbouiCOM.EditTextColumn)oAddResultGrid.Columns.Item("OrgDocNtr")).LinkedObjectType = "23";
                ((SAPbouiCOM.EditTextColumn)oAddResultGrid.Columns.Item("TrgDocNtry")).LinkedObjectType = "17";

                form.PaneLevel++;

                oAddResultGrid.AutoResizeColumns();

                //                rsInsertIntoErrorsTable.DoQuery("INSERT INTO ERRORS SELECT current_date || ' ' || current_time, 'SalesOrderSplit', '', '', 'Tran End', 'SUCCESS' FROM DUMMY;");
            }
            catch (Exception e)
            {
                Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace + "\nAt: " + sErr);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        public static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (
                pVal.FormTypeEx == "WMSalesOrderSplit" 
                && pVal.ItemUID == "ExecBtn" 
                && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                && pVal.ActionSuccess 
                && !pVal.BeforeAction
                )
            {
                New_ExecBtn_ClickAfter(null, pVal);
            }

            if (
                (                
                pVal.FormTypeEx.ToString() == "133" ||
                pVal.FormTypeEx.ToString() == "139" ||
                pVal.FormTypeEx.ToString() == "140" ||
                pVal.FormTypeEx.ToString() == "149" ||
                pVal.FormTypeEx.ToString() == "179" ||
                pVal.FormTypeEx.ToString() == "180" ||
                pVal.FormTypeEx.ToString() == "234234567"
                )
                && pVal.ItemUID == "1"
                && pVal.BeforeAction == true
                && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                )
            {
                SAPbouiCOM.Form _form = null;

                try
                {
                    _form = Application.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    //_form = Application.SBO_Application.Forms.ActiveForm;
                    if (company.UserName == "manager")
                    {
                        Recordset rsError = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        rsError.DoQuery("INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit', '', '" + company.UserName + "', 'Form: " + pVal.FormTypeEx + " " + pVal.FormTypeCount + " item: " + pVal.ItemUID + "', 'SUCCESS')");
                    }
                }
                catch (Exception e)
                {
                    Recordset rsError = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsError.DoQuery("INSERT INTO ERRORS VALUES(CURRENT_DATE || ' ' || CURRENT_TIME, 'SalesOrderSplit', '', '" + company.UserName + "', 'Form: " + pVal.FormTypeEx + " " + e.Message.Replace("'", "") + " " + e.StackTrace.ToString().Replace("'", "") + "', 'ERROR')");
                    return;
                }

                if (_form.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    && _form.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    return;
                }

                Application.SBO_Application.StatusBar.SetText("Υπολογισμός Μεικτών Εκπτώσεων", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                
                try
                {
                    _form.Freeze(true);

                    string sFinalItemCodes = "",
                       sCardCode = "",
                       sItem = "",
                       sDBDS = "",
                       sFormType = _form.Type.ToString(); // "$[CURRENT_FORMTYPE]";

                    double dQty = 0.0,
                           dDisc = 0.0;

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

                    sCardCode = _form.DataSources.DBDataSources.Item(sDBDS).GetValue("CardCode", 0).ToString();

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_form.Items.Item("38").Specific;

                    if (_form.DataSources.DBDataSources.Item(sDBDS).GetValue("DocStatus", 0).ToString() == "C"
                        && sDBDS != "OQUT")
                    {
                        return;
                    }

                    for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                    {
                        sItem = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value.ToString();
                        dQty = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific).Value, CultureInfo.InvariantCulture);

                        if (string.IsNullOrEmpty(sItem))
                        {
                            continue;
                        }

                        if (!sItem.StartsWith("ΠΦΣ"))
                        {
                            dDisc = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_TKA_SalesDiscount").Cells.Item(i).Specific).Value, CultureInfo.InvariantCulture);

                            if (dDisc != 100)
                            {
                                if (!ListItemCodes.ContainsKey(sItem))
                                {
                                    ListItemCodes.Add(sItem, dQty);

                                    //sFinalItemCodes += ";" + sItem + "?" + dQty;
                                }
                                else
                                {
                                    ListItemCodes[sItem] = ListItemCodes[sItem] + dQty;
                                }
                            }
                        }
                    }

                    foreach (KeyValuePair<string, double> sz in ListItemCodes)
                    {
                        sFinalItemCodes += ";" + sz.Key + "?" + sz.Value;
                    }

                    if (sFinalItemCodes.Length > 0)
                    {
                        sFinalItemCodes = sFinalItemCodes.Substring(1);
                    }
                    else
                    {
                        return;
                    }

                    rsPrama = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    string sSQL = "CALL TKA_SP_GET_MEIKTES_DISCOUNTS_V3('" + sFinalItemCodes + "', '" + sCardCode + "', CURRENT_DATE, '', '')";

                    rsPrama.DoQuery(sSQL);

                    if (rsPrama.RecordCount > 0)
                    {
                        if (Application.SBO_Application.MessageBox("Βρέθηκαν τιμές σετ.\nΕνημέρωση γραμμών;", 2, "Ναι", "Όχι") != 1)
                        {
                            return;
                        }
                    }

                    rsPrama.MoveFirst();

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

                                //continue;
                            }
                        }

                        rsPrama.MoveNext();
                    }
                }
                catch (Exception e)
                {
                    Application.SBO_Application.MessageBox("The Following Error Occurred:\n" + e.Message + "\n" + e.StackTrace);
                }
                finally
                {
                    _form.Freeze(false);
                }
            }
        }
    }
}