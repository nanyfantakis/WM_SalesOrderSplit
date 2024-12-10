using System;
using System.Collections.Generic;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace WM_SalesOrderSplit
{
    public class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
        //public static SAPbouiCOM.EventFilters oFilters;
        //public static SAPbouiCOM.EventFilter oFilter;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                SetFilters();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);

                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        static void SetFilters()
        {
            // Create a new EventFilters object
            /*oFilters = new SAPbouiCOM.EventFilters();
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
            oFilter.AddEx("133");
            oFilter.AddEx("139");
            oFilter.AddEx("140");
            oFilter.AddEx("149");
            oFilter.AddEx("179");
            oFilter.AddEx("180");
            oFilter.AddEx("234234567");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            oFilter.AddEx("133");
            oFilter.AddEx("139");
            oFilter.AddEx("140");
            oFilter.AddEx("149");
            oFilter.AddEx("179");
            oFilter.AddEx("180");
            oFilter.AddEx("234234567");

            Application.SBO_Application.SetFilter(oFilters);
            */
            Initialize_Event();
        }
        static void Initialize_Event()
        {
            Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            //Application.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(MeiktesEkptwseis.SBO_Application_FormDataEvent);
            //Application.SBO_Application.ItemEvent += MeiktesEkptwseis.SBO_Application_ItemEvent;
            Application.SBO_Application.ItemEvent += Menu.SBO_Application_ItemEvent;
        }


        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

    }
}
