using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace Approval
{
    class Program
    {

        private static ItemMasterApprovalService _itemMasterApprovalService;

        private static Menu _menu;

        public static SAPbobsCOM.Company oCompany { get; private set; }


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
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

                // INIT DATABASE
                DatabaseInitializer.Init();

                _menu = new Menu();
                _menu.AddMenuItems();

                oApp.RegisterMenuEventHandler(
                    _menu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);


                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();


                // INITIALIZE ITEM MASTER APPROVAL SERVICE
                // =================================================

                _itemMasterApprovalService =
                    new ItemMasterApprovalService();

                // REGISTER ITEM EVENT
                // =================================================

                Application.SBO_Application.ItemEvent +=
                    new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(
                        SBO_Application_ItemEvent);

                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        // ITEM EVENT
        // =========================================================

        //   static void SBO_Application_ItemEvent(
        //string FormUID,
        //ref SAPbouiCOM.ItemEvent pVal,
        //out bool BubbleEvent)
        //   {
        //       BubbleEvent = true;

        //       try
        //       {


        //           if (_itemMasterApprovalService == null)
        //           {
        //               Application.SBO_Application.StatusBar.SetText(
        //                   "ItemMasterApprovalService = NULL",
        //                   SAPbouiCOM.BoMessageTime.bmt_Short,
        //                   SAPbouiCOM.BoStatusBarMessageType.smt_Error);

        //               return;
        //           }

        //           _itemMasterApprovalService.HandleItemEvent(
        //               FormUID,
        //               ref pVal,
        //               out BubbleEvent);


        //           if (_approvalTemplateForm != null &&
        //               pVal.FormTypeEx == "Approval.frmApprovalTemplate")
        //           {
        //               string templateUID = "";

        //               try
        //               {
        //                   templateUID =
        //                       _approvalTemplateForm.UIAPIRawForm.UniqueID;
        //               }
        //               catch
        //               {
        //                   // Form mungkin sudah tidak bisa diakses.
        //               }

        //               if (!string.IsNullOrEmpty(templateUID) &&
        //                   pVal.FormUID == templateUID &&
        //                   pVal.EventType ==
        //                       SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
        //               {
        //                   if (pVal.BeforeAction)
        //                   {
        //                       // Tandai sebelum form benar-benar ditutup.
        //                       _approvalTemplateForm.MarkFormClosing();
        //                   }
        //                   else
        //                   {
        //                       // Lepaskan referensi setelah form ditutup.
        //                       _approvalTemplateForm = null;
        //                   }
        //               }
        //           }




        //       }
        //       catch (Exception ex)
        //       {
        //           Application.SBO_Application.MessageBox(
        //               "Error Item Master Approval Event:\n\n" +
        //               ex.ToString());
        //       }
        //   }

        static void SBO_Application_ItemEvent(
            string FormUID,
            ref SAPbouiCOM.ItemEvent pVal,
            out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Bersihkan instance Approval Template saat form ditutup.
                // Gunakan instance Menu yang sama dengan Main().
                if (_menu != null)
                {
                    _menu.HandleItemEvent(FormUID, ref pVal);
                }

                // Event Item Master tetap diproses seperti sebelumnya.
                if (_itemMasterApprovalService != null)
                {
                    _itemMasterApprovalService.HandleItemEvent(
                        FormUID,
                        ref pVal,
                        out BubbleEvent);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error Item Event: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
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
                    try
                    {
                        oCompany =
                            (SAPbobsCOM.Company)
                            Application.SBO_Application
                                .Company
                                .GetDICompany();

                        _itemMasterApprovalService =
                            new ItemMasterApprovalService();
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "Gagal reconnect ke Company:\n\n" +
                            ex.Message);
                    }


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
