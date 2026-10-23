using System;
using SAPbouiCOM.Framework;

namespace Approval
{
    class Menu
    {
        private frmApprovalTemplate _approvalTemplateForm;
        private string _approvalTemplateUID = string.Empty;


        // ==========================================================
        // ADD MENU
        // ==========================================================

        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            try
            {
                oMenus = Application.SBO_Application.Menus;

                // =================================================
                // MODULES
                // =================================================

                oMenuItem = oMenus.Item("43520");

                SAPbouiCOM.MenuCreationParams oCreationPackage =
                    (SAPbouiCOM.MenuCreationParams)
                    Application.SBO_Application.CreateObject(
                        SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                // =================================================
                // ROOT MENU
                // =================================================

                oCreationPackage.Type =
                    SAPbouiCOM.BoMenuType.mt_POPUP;

                oCreationPackage.UniqueID = "Approval";
                oCreationPackage.String = "Approval";
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = -1;

                oMenus = oMenuItem.SubMenus;

                try
                {
                    oMenus.AddEx(oCreationPackage);
                }
                catch
                {
                    // Menu sudah ada
                }


                // =================================================
                // APPROVAL MENU
                // =================================================

                oMenuItem =
                    Application.SBO_Application.Menus.Item("Approval");

                oMenus = oMenuItem.SubMenus;

                // Posisi 0: Approval Request
                AddSubMenu(
                    oMenus,
                    "Approval.Approval",
                    "Approval Request",
                    0);

                // Posisi 1: Approval Templates Setup
                AddSubMenu(
                    oMenus,
                    "Approval.frmApprovalTemplate",
                    "Approval Templates Setup",
                    1);

                // Posisi 2: Testing Only
                //AddSubMenu(
                //    oMenus,
                //    "Approval.testCheckBox",
                //    "Testing Only",
                //    2);


                
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "Gagal membuat menu Approval:\n\n" +
                    ex.ToString());
            }
        }


        // ==========================================================
        // ADD SUB MENU
        // ==========================================================


        private void AddSubMenu(
            SAPbouiCOM.Menus oMenus,
            string uniqueID,
            string caption,
            int position)
        {
            try
            {
                // Periksa apakah menu sudah ada.
                try
                {
                    SAPbouiCOM.MenuItem existingMenu =
                        oMenus.Item(uniqueID);

                    // Menu sudah ada, tidak perlu dibuat ulang.
                    return;
                }
                catch
                {
                    // Menu belum ada, lanjutkan pembuatan.
                }

                SAPbouiCOM.MenuCreationParams oCreationPackage =
                    (SAPbouiCOM.MenuCreationParams)
                    Application.SBO_Application.CreateObject(
                        SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                oCreationPackage.Type =
                    SAPbouiCOM.BoMenuType.mt_STRING;

                oCreationPackage.UniqueID = uniqueID;
                oCreationPackage.String = caption;
                oCreationPackage.Enabled = true;

                // Tentukan posisi saat membuat menu.
                oCreationPackage.Position = position;

                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Gagal membuat menu " + caption + ": " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }



        private bool IsApprovalTemplateActive()
        {
            try
            {
                if (_approvalTemplateForm == null ||
                    string.IsNullOrEmpty(_approvalTemplateUID))
                {
                    return false;
                }

                SAPbouiCOM.Form activeForm =
                    Application.SBO_Application.Forms.ActiveForm;

                if (activeForm == null)
                    return false;

                // Hanya form Approval Template yang boleh diproses.
                return activeForm.UniqueID == _approvalTemplateUID;
            }
            catch
            {
                return false;
            }
        }

        public void HandleItemEvent(
    string formUID,
    ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (_approvalTemplateForm == null ||
                    string.IsNullOrEmpty(_approvalTemplateUID))
                {
                    return;
                }

                // Pastikan event berasal dari instance Approval Template
                // yang kita buka, bukan form lain.
                if (formUID != _approvalTemplateUID)
                    return;

                if (pVal.EventType ==
                        SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD &&
                    !pVal.BeforeAction)
                {
                    _approvalTemplateForm = null;
                    _approvalTemplateUID = string.Empty;
                }
            }
            catch
            {
                // Jangan sampai cleanup mengganggu event SAP lainnya.
            }
        }

        // ==========================================================
        // MENU EVENT
        // ==========================================================

        public void SBO_Application_MenuEvent(
            ref SAPbouiCOM.MenuEvent pVal,
            out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // =================================================
                // APPROVAL TEMPLATE NAVIGATION
                // =================================================

                if (pVal.BeforeAction &&
                    (pVal.MenuUID == "1282" ||
                     pVal.MenuUID == "1288" ||
                     pVal.MenuUID == "1289" ||
                     pVal.MenuUID == "1290" ||
                     pVal.MenuUID == "1291"))
                {
                    // Jangan mencegat navigasi form SAP lain.
                    if (!IsApprovalTemplateActive())
                        return;

                    try
                    {
                        _approvalTemplateForm.HandleNavigation(pVal.MenuUID);

                        // Batalkan navigasi standar SAP hanya jika
                        // Approval Template memang sedang aktif.
                        BubbleEvent = false;
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Navigasi Approval Template gagal: " + ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }

                    return;
                }

                // =================================================
                // OPEN APPROVAL TEMPLATE
                // =================================================

                if (pVal.BeforeAction &&
                    pVal.MenuUID == "Approval.frmApprovalTemplate")
                {
                    // Jika instance lama masih valid, gunakan kembali.
                    if (_approvalTemplateForm != null)
                    {
                        try
                        {
                            SAPbouiCOM.Form existingForm =
                                Application.SBO_Application.Forms.Item(
                                    _approvalTemplateUID);

                            existingForm.Select();
                            return;
                        }
                        catch
                        {
                            // Form lama sudah tidak tersedia.
                            _approvalTemplateForm = null;
                            _approvalTemplateUID = string.Empty;
                        }
                    }

                    // Buat form baru.
                    _approvalTemplateForm = new frmApprovalTemplate();
                    _approvalTemplateForm.Show();

                    // Simpan UID form yang benar-benar dibuka.
                    _approvalTemplateUID =
                        _approvalTemplateForm.UIAPIRawForm.UniqueID;

                    return;
                }

                // =================================================
                // APPROVAL PAGE
                // =================================================

                if (pVal.BeforeAction &&
                    pVal.MenuUID == "Approval.Approval")
                {
                    Approval activeForm =
                        new Approval();

                    activeForm.Show();

                    return;
                }

                if (pVal.BeforeAction &&
                   pVal.MenuUID == "Approval.testCheckBox")
                {
                   testCheckBox activeForm =
                        new testCheckBox();

                    activeForm.Show();

                    return;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "Menu Event Error:\n\n" +
                    ex.ToString());
            }
        }
    }
}