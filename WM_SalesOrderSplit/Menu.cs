using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
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

                Application.SBO_Application.StatusBar.SetText("Μαζική Έκδοση Παραστατικών Add-On Connected Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
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
                    string sFormAsXML = Stringia.sWM_SalesOrderSplit;

                    SAPbouiCOM.FormCreationParams oCreationParams = null;

                    oCreationParams = (SAPbouiCOM.FormCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    oCreationParams.XmlData = sFormAsXML;

                    form = Application.SBO_Application.Forms.AddEx(oCreationParams);

                    Recordset rsGetPrama = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    form.DataSources.UserDataSources.Item("SlctAllPn1").Value = "Y";
                    form.DataSources.UserDataSources.Item("ElleipsiDS").Value = "N";

                    rsGetPrama = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    rsGetPrama.DoQuery("SELECT \"Series\", \"SeriesName\" FROM NNM1 WHERE \"ObjectCode\" = '17' AND \"SeriesName\" != 'ΕΠΙΛΕΞΤΕ'");

                    SAPbouiCOM.ComboBox oCB1 = (SAPbouiCOM.ComboBox)form.Items.Item("SeriesTo").Specific;

                    while (!rsGetPrama.EoF)
                    {
                        oCB1.ValidValues.Add(rsGetPrama.Fields.Item(0).Value.ToString(), rsGetPrama.Fields.Item(1).Value.ToString());

                        rsGetPrama.MoveNext();
                    }

                    /*rsGetPrama.DoQuery("SELECT \"Series\", \"SeriesName\" FROM NNM1 WHERE \"ObjectCode\" = '23' " +
                                        " UNION ALL SELECT '-1', '' FROM DUMMY ORDER BY 1 ");

                    oCB1 = (SAPbouiCOM.ComboBox)form.Items.Item("SeriesFrom").Specific;

                    while (!rsGetPrama.EoF)
                    {
                        oCB1.ValidValues.Add(rsGetPrama.Fields.Item(0).Value.ToString(), rsGetPrama.Fields.Item(1).Value.ToString());

                        rsGetPrama.MoveNext();
                    }*/

                    rsGetPrama.DoQuery("SELECT T0.\"Name\", T0.\"Code\" FROM \"@TKA_SUBCATEGORY\" T0 " +
                                        " UNION ALL SELECT '', '' FROM DUMMY ORDER BY 1");

                    oCB1 = (SAPbouiCOM.ComboBox)form.Items.Item("SubCatCB").Specific;

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

                    Initialize_Form();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        public void Initialize_Form()
        {
            try
            {
                form = Application.SBO_Application.Forms.ActiveForm;

                Users oUser = (Users)company.GetBusinessObject(BoObjectTypes.oUsers);

                oUser.GetByKey(((ICompany)Application.SBO_Application.Company.GetDICompany()).UserSignature);

                form.DataSources.UserDataSources.Item("SlctAllPn1").Value = "Y";

                ((SAPbouiCOM.Button)form.Items.Item("SelectBTN").Specific).ClickAfter += SelectBTN_ClickAfter;
                ((SAPbouiCOM.Button)form.Items.Item("BackBtn").Specific).ClickAfter += BackBtn_ClickAfter;
                ((SAPbouiCOM.Button)form.Items.Item("NextBtn").Specific).ClickAfter += NextBtn_ClickAfter;
                ((SAPbouiCOM.Button)form.Items.Item("ExecBtn").Specific).ClickAfter += ExecBtn_ClickAfter;
                ((SAPbouiCOM.Button)form.Items.Item("RefreshBT").Specific).ClickAfter += NextBtn_ClickAfter;

                ((SAPbouiCOM.EditText)form.Items.Item("CardCodeED").Specific).ValidateAfter += CardCodeED_ValidateAfter;
                ((SAPbouiCOM.EditText)form.Items.Item("CardCodeED").Specific).ChooseFromListAfter += CardCodeED_ChooseFromListAfter;

                ((SAPbouiCOM.Grid)form.Items.Item("AddResGRD").Specific).Columns.Item("OrgDocNtr").LinkPressedBefore += AddResGRD_OrgDocNtr_LinkPressedBefore;
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
                        oGrid.DataTable.SetValue("Check", oGrid.GetDataTableRowIndex(i), "Y");
                    }
                    form.DataSources.UserDataSources.Item("SlctAllPn1").Value = "N";
                }
                else
                {
                    for (int i = 0; i < oGrid.Rows.Count; i++)
                    {
                        oGrid.DataTable.SetValue("Check", oGrid.GetDataTableRowIndex(i), "N");
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

        private void AddResGRD_OrgDocNtr_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            string sDocEntry = ((SAPbouiCOM.EditText)((SAPbouiCOM.Grid)form.Items.Item("AddResGRD").Specific).Columns.Item("OrgDocNtr")).Value.ToString();

            int iObjType = Convert.ToInt32(((SAPbouiCOM.EditText)((SAPbouiCOM.Grid)form.Items.Item("AddResGRD").Specific).Columns.Item("OriginType")).Value.ToString(), System.Globalization.CultureInfo.InvariantCulture);

            Application.SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)iObjType, "", sDocEntry);

            BubbleEvent = false;
        }

        private void BackBtn_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                form.PaneLevel--;

                if (form.PaneLevel == 2)
                {
                    form.Items.Item("RefreshBT").Click();
                }
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
                   sCardCode;

            SAPbouiCOM.Grid oFinalGrid;

            try
            {
                if (form.PaneLevel == 1)
                {
                    if (string.IsNullOrEmpty(form.DataSources.UserDataSources.Item("SeriesTo").Value.ToString()))
                    {
                        Application.SBO_Application.StatusBar.SetText("Παρακαλώ Συμπληρώστε Σειρά Παραστατικού Στόχου.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        form.ActiveItem = "SeriesTo";
                        return;
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

                form.Freeze(true);
                if (form.PaneLevel != 2)
                {
                    form.PaneLevel++;
                }

                if (form.PaneLevel == 2)
                {
                    sErr = "Build sSQL";
//                    sSQL = string.Format(Stringia.sNextSQL, form.DataSources.UserDataSources.Item("DocDateDS").Value.ToString(), form.DataSources.UserDataSources.Item("ElleipsiDS").Value.ToString());
                    sSQL = string.Format(Stringia.sNextSQL, form.DataSources.UserDataSources.Item("ElleipsiDS").Value.ToString());

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
                        sSQL += "Q.\"CardCode\" = '" + sCardCode + "' ";
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
                            "   Q.\"NumAtCard\" ";

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


                    for (int i = 0; i < oFinalGrid.DataTable.Rows.Count; i++)
                    {
                        if (oFinalGrid.DataTable.GetValue("ZERO_IS_INVALID", i).ToString() == "0")
                        {
                            oFinalGrid.CommonSetting.SetRowEditable(i+1, false);
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

        private void ExecBtn_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string sDocEntries = "0",
                   sErr = "",
                   sSQL,
                   sCardCode,
                   sSQLLines = "SELECT \"AliasID\" FROM CUFD WHERE \"TableID\" = 'RDR1'",
                   sSQLHeader = "SELECT \"AliasID\" FROM CUFD WHERE \"TableID\" = 'ORDR'";

            SAPbouiCOM.Grid oGrid,
                            oAddResultGrid;

            SAPbouiCOM.DataTable oResultDT;

            bool bSelected = false;

            Recordset rsGetData,
                      rsInsertIntoErrorsTable;

            Documents oOrder,
                      oBaseDoc = (Documents)company.GetBusinessObject(BoObjectTypes.oQuotations); ;

            int iDocEntry,
                iProperty;

            DateTime dtDocDate;

            try
            {
                if (Application.SBO_Application.MessageBox("Δημιουργία Παραγγελιών για τις επιλεγμένες γραμμές;", 2, "Ναι", "Όχι") != 1)
                {
                    return;
                }
                rsInsertIntoErrorsTable = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsInsertIntoErrorsTable.DoQuery("INSERT INTO ERRORS SELECT current_date || ' ' || current_time, 'SalesOrderSplit', '', '', 'Tran Start', 'SUCCESS' FROM DUMMY;");

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
                }

                sSQL = string.Format(Stringia.sExecSQL, sDocEntries);

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
                rsDateTimeFormat = null;
                oSBObob = null;

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
                        }
                        else
                        {
                            oResultDT.Rows.Add();

                            oResultDT.SetValue("AddRes", oResultDT.Rows.Count - 1, "Επιτυχία");
                            oResultDT.SetValue("OrgDocNtr", oResultDT.Rows.Count - 1, oBaseDoc.DocEntry);
                            oResultDT.SetValue("OriginType", oResultDT.Rows.Count - 1, "23");
                            oResultDT.SetValue("TrgDocNtry", oResultDT.Rows.Count - 1, company.GetNewObjectKey());
                            oResultDT.SetValue("TargetType", oResultDT.Rows.Count - 1, "17");
                            oResultDT.SetValue("SapErrCode", oResultDT.Rows.Count - 1, -1);
                            oResultDT.SetValue("SapErrMsg", oResultDT.Rows.Count - 1, string.Empty);
                        }

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

                    Recordset oRSFieldsHeader = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    oRSFieldsHeader.DoQuery(sSQLHeader);
                    oRSFieldsHeader.MoveFirst();

                    while (!oRSFieldsHeader.EoF)
                    {
                        string sFieldName = "U_" + oRSFieldsHeader.Fields.Item("AliasID").Value.ToString() + "";

                        if (!string.IsNullOrEmpty(oBaseDoc.UserFields.Fields.Item(sFieldName).Value.ToString()))
                        {
                            oOrder.UserFields.Fields.Item(sFieldName).Value = oBaseDoc.UserFields.Fields.Item(sFieldName).Value;
                        }

                        oRSFieldsHeader.MoveNext();
                    }

                    oBaseDoc.Lines.SetCurrentLine(Convert.ToInt32(rsGetData.Fields.Item("LineNum").Value));

                    oOrder.Lines.BaseLine = Convert.ToInt32(rsGetData.Fields.Item("LineNum").Value);
                    oOrder.Lines.BaseEntry = iDocEntry;
                    oOrder.Lines.BaseType = 23;

                    sErr = "Get U_Fields for Lines - Final Doc";
                    Recordset oRSFields = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRSFields.DoQuery(sSQLLines);
                    oRSFields.MoveFirst();

                    while (oRSFields.EoF == false)
                    {
                        string sFieldName = "U_" + oRSFields.Fields.Item("AliasID").Value.ToString() + "";

                        if (!string.IsNullOrEmpty(oBaseDoc.Lines.UserFields.Fields.Item(sFieldName).Value.ToString()))
                        {
                            oOrder.Lines.UserFields.Fields.Item(sFieldName).Value = oBaseDoc.Lines.UserFields.Fields.Item(sFieldName).Value;
                        }

                        oRSFields.MoveNext();
                    }

                    oOrder.Lines.Add();

                    prg.Value++;

                    rsGetData.MoveNext();
                }

                try
                {
                    prg.Stop();
                }
                catch (Exception)
                { }

                prg = null;

                ((SAPbouiCOM.EditTextColumn)oAddResultGrid.Columns.Item("OrgDocNtr")).LinkedObjectType = "23";
                ((SAPbouiCOM.EditTextColumn)oAddResultGrid.Columns.Item("TrgDocNtry")).LinkedObjectType = "17";

                form.PaneLevel++;

                oAddResultGrid.AutoResizeColumns();

                rsInsertIntoErrorsTable.DoQuery("INSERT INTO ERRORS SELECT current_date || ' ' || current_time, 'SalesOrderSplit', '', '', 'Tran End', 'SUCCESS' FROM DUMMY;");
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
    }
}
