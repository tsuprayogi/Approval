using System;
using System.Runtime.InteropServices;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace Approval
{
    public class ItemMasterApprovalService
    {
        private SAPbobsCOM.Company _company;

        private string _pendingItemCode;
        private string _pendingItemName;

        // Mencegah event diproses dua kali
        private bool _processingAdd = false;

        public ItemMasterApprovalService()
        {
            _company =
                (SAPbobsCOM.Company)
                Application.SBO_Application.Company.GetDICompany();
        }

        // =========================================================
        // DYNAMIC OBJECT CONFIGURATION CACHE (@SBO_OBJECT)
        // =========================================================

        private sealed class SsbObjectConfig
        {
            public string Code { get; set; }            
            public int ObjType { get; set; }
            public string BaseType { get; set; }
            public string ObjName { get; set; }
            public string ObjTable { get; set; }
            public int FormId { get; set; }
            
        }

        private static readonly object _objectConfigLock = new object();
        private static System.Collections.Generic.Dictionary<int, SsbObjectConfig>
            _objectConfigCache = null;

        private string _pendingFormUID;

        private static void LoadObjectConfig(bool forceReload = false)
        {
            lock (_objectConfigLock)
            {
                if (_objectConfigCache != null && !forceReload)
                    return;

                var configs = new System.Collections.Generic.Dictionary<int, SsbObjectConfig>();
                SAPbobsCOM.Recordset rs = null;

                try
                {
                    SAPbobsCOM.Company company =
                        (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                    rs = (SAPbobsCOM.Recordset)company.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sql = @"
                        SELECT [Code], [Name]  ,[U_T2_ObjGroup]
                             ,[U_T2_ObjCode]
                             ,[U_T2_ObjBase]
                             ,[U_T2_ObjName]
                             ,[U_T2_ObjTable]
                             ,[U_T2_FormId]
                        FROM [@SBO_OBJECT]
                        WHERE ISNULL([U_T2_FormId], '') <> ''
                          AND TRY_CONVERT(INT, [U_T2_FormId]) IS NOT NULL";

                    rs.DoQuery(sql);

                    while (!rs.EoF)
                    {
                        int formId = Convert.ToInt32(rs.Fields.Item("U_T2_FormId").Value);
                        int objType;
                        int.TryParse(Convert.ToString(rs.Fields.Item("U_T2_ObjCode").Value), out objType);

                        var config = new SsbObjectConfig
                        {
                            Code = Convert.ToString(rs.Fields.Item("Code").Value).Trim(),
                            ObjTable = Convert.ToString(rs.Fields.Item("U_T2_ObjTable").Value).Trim(),
                            ObjType = objType,
                            BaseType = Convert.ToString(rs.Fields.Item("U_T2_ObjBase").Value).Trim(),
                            FormId = formId,
                            ObjName = Convert.ToString(rs.Fields.Item("U_T2_ObjName").Value).Trim()
                        };

                        if (!configs.ContainsKey(formId))
                        {
                            configs.Add(formId, config);
                        }
                        else
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                "Peringatan: U_IDForm duplikat " + formId +
                                " pada @SBO_OBJECT. Form ini hanya memakai konfigurasi pertama.",
                                BoMessageTime.bmt_Short,
                                BoStatusBarMessageType.smt_Warning);
                        }

                        rs.MoveNext();
                    }

                    _objectConfigCache = configs;
                }
                finally
                {
                    if (rs != null)
                        Marshal.ReleaseComObject(rs);
                }
            }
        }

        private static bool TryGetObjectConfig(
            string formTypeEx,
            out SsbObjectConfig config)
        {
            config = null;

            int formId;
            if (!int.TryParse(formTypeEx, out formId))
                return false;

            LoadObjectConfig();

            lock (_objectConfigLock)
            {
                return _objectConfigCache.TryGetValue(formId, out config);
            }
        }

        // Panggil setelah data @SBO_OBJECT berhasil ditambah/diubah.
        public static void RefreshObjectConfig()
        {
            LoadObjectConfig(forceReload: true);
        }

        // =========================================================
        // EVENT HANDLER
        // =========================================================

        public void HandleItemEvent(
            string formUID,
            ref ItemEvent pVal,
            out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {

                // Cari konfigurasi form dari @SBO_OBJECT.
                SsbObjectConfig objectConfig;

                if (!TryGetObjectConfig(pVal.FormTypeEx, out objectConfig))
                    return;

                // Validasi object berdasarkan konfigurasi @SBO_OBJECT.
                string ObjTable = (objectConfig.ObjTable ?? string.Empty)
                    .Trim();

                bool isItemMaster =
                    objectConfig.ObjType == 4 &&
                    string.Equals(
                        ObjTable,
                        "OITM",
                        StringComparison.OrdinalIgnoreCase);

                bool isBusinessPartner =
                    objectConfig.ObjType == 2 &&
                    string.Equals(
                        ObjTable,
                        "OCRD",
                        StringComparison.OrdinalIgnoreCase);

                // Proses hanya OITM dan OCRD.
                if (!isItemMaster && !isBusinessPartner)
                    return;

                // Hanya event klik tombol Add SAP.
                if (pVal.EventType != BoEventTypes.et_ITEM_PRESSED ||
                    pVal.ItemUID != "1")
                    return;

                // ---------------------------------------------------------
                // BEFORE ACTION: simpan data sebelum SAP melakukan Add.
                // ---------------------------------------------------------
                if (pVal.BeforeAction)
                {
                    _pendingItemCode = string.Empty;
                    _pendingItemName = string.Empty;
                    _pendingFormUID = string.Empty;

                    try
                    {
                        SAPbouiCOM.Form form =
                            Application.SBO_Application.Forms.Item(formUID);

                        if (form.Mode != BoFormMode.fm_ADD_MODE)
                            return;

                        GetItemDataFromForm(
                            formUID,
                            out _pendingItemCode,
                            out _pendingItemName);

                        _pendingFormUID = formUID;

                        Application.SBO_Application.StatusBar.SetText(
                            "Persiapan Add Item: " + _pendingItemCode,
                            BoMessageTime.bmt_Short,
                            BoStatusBarMessageType.smt_Warning);
                    }
                    catch (Exception ex)
                    {
                        ClearPendingItem();
                        Application.SBO_Application.StatusBar.SetText(
                            "Gagal membaca Item Master: " + ex.Message,
                            BoMessageTime.bmt_Short,
                            BoStatusBarMessageType.smt_Error);
                    }

                    return;
                }

                // ---------------------------------------------------------
                // AFTER ACTION: proses hanya jika SAP benar-benar berhasil Add.
                // ---------------------------------------------------------
                if (!pVal.ActionSuccess)
                {
                    if (string.Equals(_pendingFormUID, formUID, StringComparison.Ordinal))
                        ClearPendingItem();
                    return;
                }

                // Pastikan data pending berasal dari form yang sama.
                if (!string.Equals(_pendingFormUID, formUID, StringComparison.Ordinal) ||
                    string.IsNullOrWhiteSpace(_pendingItemCode) ||
                    _pendingItemCode == "*")
                {
                    ClearPendingItem();
                    return;
                }

                if (_processingAdd)
                    return;

                _processingAdd = true;
              
                try
                {
                    
                    switch (ObjTable)
                    {
                        case "OITM":
                            ProcessNewItem(
                                _pendingItemCode,
                                _pendingItemName);
                            break;

                        case "OCRD":
                            ProcessNewBP(
                                _pendingItemCode,
                                _pendingItemName);
                            break;

                        default:
                            Application.SBO_Application.StatusBar.SetText(
                                "Object belum memiliki handler approval: " +
                                ObjTable,
                                BoMessageTime.bmt_Short,
                                BoStatusBarMessageType.smt_Warning);
                            break;
                    }

                }
                finally
                {
                    ClearPendingItem();
                    _processingAdd = false;
                }
            }
            catch (Exception ex)
            {
                ClearPendingItem();
                Application.SBO_Application.MessageBox(
                    "Error Item Master Approval:\n" + ex.ToString());
            }
        }

        private void ClearPendingItem()
        {
            _pendingItemCode = string.Empty;
            _pendingItemName = string.Empty;
            _pendingFormUID = string.Empty;
        }

        private void GetItemDataFromForm(
                   string formUID,
                   out string itemCode,
                   out string itemName)
        {
            SAPbouiCOM.Form oForm = null;

            itemCode = string.Empty;
            itemName = string.Empty;

            try
            {
                oForm =
                    Application.SBO_Application.Forms.Item(formUID);


                // =================================================
                // ITEM CODE
                // FORM 150 - ITEM UID 5
                // =================================================
                SAPbouiCOM.EditText txtItemCode =
                    (SAPbouiCOM.EditText)
                    oForm.Items.Item("5").Specific;

                itemCode =
                    txtItemCode.Value.Trim();


                // =================================================
                // ITEM NAME
                // FORM 150 - ITEM UID 7
                // =================================================
                SAPbouiCOM.EditText txtItemName =
                    (SAPbouiCOM.EditText)
                    oForm.Items.Item("7").Specific;

                itemName =
                    txtItemName.Value.Trim();
            }
            finally
            {
                if (oForm != null)
                {
                    System.Runtime.InteropServices.Marshal
                        .ReleaseComObject(oForm);
                }
            }
        }


        private void ProcessNewBP(string itemCode, string itemName)
        {
            SAPbobsCOM.Recordset rs = null;
            SAPbobsCOM.BusinessPartners partners = null;

            try
            {
                // =====================================================
                // 1. AMBIL DATA BP YANG SUDAH BERHASIL DIBUAT
                // =====================================================
                rs = (SAPbobsCOM.Recordset)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = @"
            SELECT
                DocEntry,
                CardCode,
                CardName,
                ObjType
            FROM OCRD
            WHERE CardCode = '" + SqlSafe(itemCode) + "'";

                rs.DoQuery(sql);

                if (rs.EoF)
                {
                    throw new Exception(
                        "Business Partner tidak ditemukan di OCRD.\n" +
                        "BPCode: " + itemCode);
                }

                int docEntry = Convert.ToInt32(
                    rs.Fields.Item("DocEntry").Value);

                string dbItemCode = Convert.ToString(
                    rs.Fields.Item("CardCode").Value).Trim();

                string dbItemName = Convert.ToString(
                    rs.Fields.Item("CardName").Value).Trim();

                string dbObjType = Convert.ToString(
                    rs.Fields.Item("ObjType").Value).Trim();

                // =====================================================
                // 2. AMBIL USER LOGIN SAP
                // =====================================================
                string userCode = GetCurrentUserCode();

                if (string.IsNullOrWhiteSpace(userCode))
                {
                    throw new Exception(
                        "User login SAP tidak ditemukan.");
                }

                // Validasi konfigurasi approval sebelum mengubah BP
                ItemApprovalConfig approvalConfig =
                GetApprovalConfig(
                    userCode,
                    dbObjType,
                    dbItemCode,
                    "OCRD");

                if (approvalConfig == null)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "BP tidak memenuhi kondisi template approval. " +
                        "Approval dilewati: " + dbItemCode,
                        BoMessageTime.bmt_Medium,
                        BoStatusBarMessageType.smt_Warning);

                    return;
                }

                // =====================================================
                // 3. AMBIL BP DAN NONAKTIFKAN
                // =====================================================
                partners = (SAPbobsCOM.BusinessPartners)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                if (!partners.GetByKey(itemCode))
                {
                    throw new Exception(
                        "Gagal mendapatkan Business Partner: " +
                        itemCode);
                }

                SetBPToApproval(itemCode, docEntry, userCode);


               

                // =====================================================
                // 4. BUAT APPROVAL HEADER
                // =====================================================
                CreateApprovalHeader(
                    docEntry,
                    dbItemCode,
                    dbItemName,
                    userCode,
                    dbObjType,
                    approvalConfig);

                // =====================================================
                // 5. BUAT APPROVAL DETAIL
                // =====================================================
                CreateApprovalDetail(
                    docEntry,
                    approvalConfig);

                // =====================================================
                // 6. NOTIFIKASI BERHASIL
                // =====================================================
                Application.SBO_Application.StatusBar.SetText(
                    "BP berhasil dinonaktifkan dan masuk approval. " +
                    "BPCode: " + dbItemCode +
                    " | BPName: " + dbItemName,
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Proses approval BP gagal: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                throw;
            }
            finally
            {
                if (partners != null)
                {
                    System.Runtime.InteropServices.Marshal
                        .ReleaseComObject(partners);
                    partners = null;
                }

                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal
                        .ReleaseComObject(rs);
                    rs = null;
                }
            }
        }

        private void ProcessNewItem(
            string itemCode,
            string itemName)
        {
            SAPbobsCOM.Recordset rs = null;
            SAPbobsCOM.Items item = null;

            try
            {
                // =====================================================
                // 1. CARI DATA ITEM DI OITM
                // =====================================================
                rs = (SAPbobsCOM.Recordset)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = @"
                SELECT
                    DocEntry,
                    ItemCode,
                    ItemName,
                    ObjType
                FROM OITM
                WHERE ItemCode = '" +
                    SqlSafe(itemCode) +
                    "'";

                rs.DoQuery(sql);

                if (rs.EoF)
                {
                    Application.SBO_Application.MessageBox(
                        "Item tidak ditemukan di OITM.\n\n" +
                        "ItemCode : " + itemCode);

                    return;
                }

                int docEntry =
                    Convert.ToInt32(
                        rs.Fields.Item("DocEntry").Value);

                string dbItemCode =
                    Convert.ToString(
                        rs.Fields.Item("ItemCode").Value).Trim();

                string dbItemName =
                    Convert.ToString(
                        rs.Fields.Item("ItemName").Value).Trim();

                string dbObjType =
                    Convert.ToString(
                        rs.Fields.Item("ObjType").Value).Trim();



                // =====================================================
                // 2. AMBIL USER DAN VALIDASI WORKFLOW
                // =====================================================
                string userCode = GetCurrentUserCode();


                if (string.IsNullOrWhiteSpace(userCode))
                {
                    throw new Exception(
                        "User login SAP tidak ditemukan.");
                }

                // Validasi harus selesai sebelum OITM diubah
                ItemApprovalConfig approvalConfig =
                   GetApprovalConfig(
                       userCode,
                       dbObjType,
                       dbItemCode,
                       "OITM");

                if (approvalConfig == null)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Item tidak memenuhi kondisi template approval. " +
                        "Approval dilewati: " + dbItemCode,
                        BoMessageTime.bmt_Medium,
                        BoStatusBarMessageType.smt_Warning);

                    return;
                }




                // =====================================================
                // 4. UPDATE OITM
                // =====================================================
                item = (SAPbobsCOM.Items)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.oItems);

                if (!item.GetByKey(itemCode))
                {
                    throw new Exception(
                        "Gagal mendapatkan Item:\n" +
                        itemCode);
                }

                SetItemToApproval(itemCode, docEntry, userCode);
                              
                // 5. INSERT APPROVAL HEADER              
                CreateApprovalHeader(
                    docEntry,
                    dbItemCode,
                    dbItemName,
                    userCode,
                    dbObjType,
                    approvalConfig);
             
                // 6. INSERT APPROVAL DETAIL             
                CreateApprovalDetail(
                     docEntry,
                     approvalConfig);
                
                // 7. HASIL                
                Application.SBO_Application.StatusBar.SetText(
                    "ITEM BERHASIL MASUK APPROVAL\n\n" +
                    //"DocEntry : " + docEntry +
                    "\nItemCode : " + dbItemCode +
                    "\nItemName : " + dbItemName +
                    //"\nStatus   : P" +
                    //"\nValid    : NO" +
                    //"\nFrozen   : YES" +
                    "\nApproval : CREATED", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "ProcessNewItem Error:\n\n" +
                    ex.ToString());
            }
            finally
            {
                if (item != null)
                {
                    System.Runtime.InteropServices.Marshal
                        .ReleaseComObject(item);
                }

                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal
                        .ReleaseComObject(rs);
                }
            }
        }



        // GET ITEM CODE      
        //private string GetItemCodeFromForm(string formUID)
        //{
        //    SAPbouiCOM.Form form = null;

        //    try
        //    {
        //        form = Application.SBO_Application.Forms.Item(formUID);

        //        // Item Master menggunakan DBDataSource OITM
        //        SAPbouiCOM.DBDataSource ds =
        //            form.DataSources.DBDataSources.Item("OITM");

        //        return ds.GetValue("ItemCode", 0).Trim();
        //    }
        //    catch
        //    {
        //        // Fallback: ambil dari ItemUID ItemCode
        //        try
        //        {
        //            SAPbouiCOM.EditText itemCodeEdit =
        //                (SAPbouiCOM.EditText)
        //                form.Items.Item("5").Specific;

        //            return itemCodeEdit.Value.Trim();
        //        }
        //        catch
        //        {
        //            return "";
        //        }
        //    }
        //}


        // SET ITEM MENJADI PENDING APPROVAL

        private void SetItemToApproval(
            string itemCode,
            int docEntry,
            string userCode)
        {
            SAPbobsCOM.Items item = null;

            try
            {
                item = (SAPbobsCOM.Items)
                    _company.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.oItems);

                if (!item.GetByKey(itemCode))
                {
                    throw new Exception(
                        "Gagal mengambil Item: " + itemCode);
                }

                // =================================================
                // ITEM TIDAK AKTIF
                // =================================================

                item.Valid =
                    SAPbobsCOM.BoYesNoEnum.tNO;

                // =================================================
                // ITEM DI-FREEZE
                // =================================================

                item.Frozen =
                    SAPbobsCOM.BoYesNoEnum.tYES;

                // =================================================
                // UDF APPROVAL
                // =================================================

                item.UserFields.Fields
                    .Item("U_ApprovalStatus")
                    .Value = "P";

                item.UserFields.Fields
                    .Item("U_ApprovalDocEntry")
                    .Value = docEntry.ToString();

                item.UserFields.Fields
                    .Item("U_ApprovalUser")
                    .Value = userCode;

                // =================================================
                // UPDATE
                // =================================================

                int result = item.Update();

                if (result != 0)
                {
                    int errorCode;
                    string errorMessage;

                    _company.GetLastError(
                        out errorCode,
                        out errorMessage);

                    throw new Exception(
                        "Gagal mengubah status Item: " +
                        errorCode +
                        " - " +
                        errorMessage);
                }
            }
            finally
            {
                if (item != null)
                    Marshal.ReleaseComObject(item);
            }
        }


        private void SetBPToApproval(
              string CardCode,
              int docEntry,
              string userCode)
        {
            SAPbobsCOM.BusinessPartners partners = null;

            try
            {
                partners = (SAPbobsCOM.BusinessPartners)
                    _company.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                if (!partners.GetByKey(CardCode))
                {
                    throw new Exception(
                        "Gagal mengambil Item: " + CardCode);
                }
              
                // =================================================
                // ITEM TIDAK AKTIF
                // =================================================

                partners.Valid =
                    SAPbobsCOM.BoYesNoEnum.tNO;

                // =================================================
                // ITEM DI-FREEZE
                // =================================================

                partners.Frozen =
                    SAPbobsCOM.BoYesNoEnum.tYES;

                // =================================================
                // UDF APPROVAL
                // =================================================



                partners.UserFields.Fields
                    .Item("U_ApprovalStatus")
                    .Value = "P";

                partners.UserFields.Fields
                    .Item("U_ApprovalDocEntry")
                    .Value =  docEntry.ToString();
                ;

                partners.UserFields.Fields
                    .Item("U_ApprovalUser")
                    .Value = userCode;

                // =================================================
                // UPDATE
                // =================================================

                int result = partners.Update();

                if (result != 0)
                {
                    int errorCode;
                    string errorMessage;

                    _company.GetLastError(
                        out errorCode,
                        out errorMessage);

                    throw new Exception(
                        "Gagal mengubah status BP: " +
                        errorCode +
                        " - " +
                        errorMessage);
                }
            }
            finally
            {
                if (partners != null)
                    Marshal.ReleaseComObject(partners);
            }
        }


        private bool EvaluateTemplateCondition(
    string condition,
    string objTable,
    string objectCode)
        {
            // Kondisi kosong berarti tidak ada filter.
            if (string.IsNullOrWhiteSpace(condition))
                return true;

            condition = condition.Trim();

            // Hanya izinkan query SELECT.
            if (!condition.StartsWith(
                "SELECT ",
                StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception(
                    "U_TCondition harus berupa query SELECT.");
            }

            // Tolak beberapa pola SQL berisiko.
            string upperCondition = condition.ToUpperInvariant();

            if (condition.Contains(";") ||
                condition.Contains("--") ||
                condition.Contains("/*") ||
                condition.Contains("*/") ||
                System.Text.RegularExpressions.Regex.IsMatch(
                    upperCondition,
                    @"\b(INSERT|UPDATE|DELETE|DROP|ALTER|TRUNCATE|EXEC|EXECUTE)\b"))
            {
                throw new Exception(
                    "U_TCondition mengandung sintaks SQL yang tidak diizinkan.");
            }

            // Tentukan field kunci berdasarkan tabel.
            string keyField;

            if (string.Equals(
                objTable, "OCRD",
                StringComparison.OrdinalIgnoreCase))
            {
                keyField = "CardCode";
            }
            else if (string.Equals(
                objTable, "OITM",
                StringComparison.OrdinalIgnoreCase))
            {
                keyField = "ItemCode";
            }
            else
            {
                throw new Exception(
                    "Tabel tidak didukung untuk kondisi approval: " +
                    objTable);
            }

            SAPbobsCOM.Recordset rs = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)
                    _company.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                rs.DoQuery(condition);

                // U_TCondition wajib menghasilkan kolom kode objek
                // sebagai kolom pertama.
                if (rs.Fields.Count == 0)
                {
                    throw new Exception(
                        "U_TCondition tidak menghasilkan kolom kode objek.");
                }

                while (!rs.EoF)
                {
                    string resultCode = Convert.ToString(
                        rs.Fields.Item(0).Value).Trim();

                    if (string.Equals(
                        resultCode,
                        objectCode,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    rs.MoveNext();
                }

                return false;
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }



        #region Utama

        private void CreateApprovalHeader(
            int docEntry,
            string itemCode,
            string itemName,
            string userCode,
            string objtype,
            ItemApprovalConfig config)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = @"
            INSERT INTO [@AITM_APV]
            (
                Code,
                Name,
                U_DocType,
                U_ItemCode,
                U_ItemName,
                U_Status,
                U_RequestBy,
                U_RequestDate,
                U_StageID,
                U_CurrStep,
                U_StageStatus
            )
            VALUES
            (
                '" + docEntry + @"',
                '" + docEntry + @"',
                '" + SqlSafe(objtype) + @"',
                '" + SqlSafe(itemCode) + @"',
                '" + SqlSafe(itemName) + @"',
                'PENDING',
                '" + SqlSafe(userCode) + @"',
                GETDATE(),
                " + config.StageID + @",
                " + config.CurrentStep + @",
                'PENDING'
            )";

                rs.DoQuery(sql);
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }


        private void CreateApprovalDetail(
            int docEntry,
            ItemApprovalConfig config)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = @"
            INSERT INTO [@AITM_APV_L]
            (
                Code,
                Name,
                U_StageID,
                U_Approver,
                U_Status
            )
            VALUES
            (
                '" + docEntry + @"',
                '" + docEntry + @"',
                " + config.StageID + @",
                '" + SqlSafe(config.UserCode) + @"',
                'PENDING'
            )";

                rs.DoQuery(sql);
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }


        private string GetCurrentUserCode()
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                SAPbobsCOM.Company company =
                    (SAPbobsCOM.Company)
                    Application.SBO_Application.Company.GetDICompany();

                rs = (SAPbobsCOM.Recordset)
                    company.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                int userId = company.UserSignature;

                string sql = @"
            SELECT USER_CODE
            FROM OUSR
            WHERE USERID = " + userId;

                rs.DoQuery(sql);

                if (rs.EoF)
                    return string.Empty;

                return Convert.ToString(
                    rs.Fields.Item("USER_CODE").Value
                ).Trim();
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }



        private string GetField(
            SAPbobsCOM.Recordset rs,
            string fieldName)
        {
            object value =
                rs.Fields.Item(fieldName).Value;

            if (value == null ||
                value == DBNull.Value)
                return "";

            return Convert.ToString(value).Trim();
        }

        private string SqlSafe(string value)
        {
            return (value ?? "")
                .Replace("'", "''");
        }


        public class ItemApprovalConfig
        {
            public string TemplateCode { get; set; }
            public string TemplateName { get; set; }
            public string TemplateRemarks { get; set; }
            public string UserCode { get; set; }
            public string ObjType { get; set; }

            public int StageID { get; set; }
            public int CurrentStep { get; set; }
        }


        private ItemApprovalConfig GetApprovalConfig(
    string userCode,
    string objType,
    string objectCode,
    string objTable)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)
                    _company.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string safeUser = SqlSafe(userCode);
                string safeType = SqlSafe(objType);

                string sql = @"
            SELECT
                T.Code AS TemplateCode,
                T.Name AS TemplateName,
                T.U_TRemarks AS TemplateRemarks,
                T.U_TCondition AS TemplateCondition,
                W.U_U_UserID AS UserCode,
                T.U_ObjectType AS ObjType,
                S.U_U_StageCode AS StageID,
                S.U_U_SortId AS CurrentStep
            FROM [@WTM1] W
            INNER JOIN [@OWTM] T
                ON W.Code LIKE T.Code + '-%'
            INNER JOIN [@WTM2] S
                ON S.Code = W.Code
            WHERE W.U_U_UserID = '" + safeUser + @"'
              AND T.U_U_Active = 'Y'
              AND T.U_ObjectType = '" + safeType + @"'
            ORDER BY
                S.U_U_SortId ASC,
                T.Code DESC";

                rs.DoQuery(sql);

                if (rs.EoF)
                {
                    throw new Exception(
                        "Workflow approval tidak ditemukan atau tidak aktif.\n" +
                        "User: " + userCode +
                        " | Object Type: " + objType);
                }

                while (!rs.EoF)
                {
                    string templateCondition = Convert.ToString(
                        rs.Fields.Item("TemplateCondition").Value).Trim();

                    // Evaluasi kondisi template terhadap objek saat ini.
                    bool conditionPassed = EvaluateTemplateCondition(
                        templateCondition,
                        objTable,
                        objectCode);

                    if (conditionPassed)
                    {
                        ItemApprovalConfig config =
                            new ItemApprovalConfig();

                        config.TemplateCode =
                            Convert.ToString(
                                rs.Fields.Item("TemplateCode").Value).Trim();

                        config.TemplateName =
                            Convert.ToString(
                                rs.Fields.Item("TemplateName").Value).Trim();

                        config.TemplateRemarks =
                            Convert.ToString(
                                rs.Fields.Item("TemplateRemarks").Value).Trim();

                        config.UserCode =
                            Convert.ToString(
                                rs.Fields.Item("UserCode").Value).Trim();

                        config.ObjType =
                            Convert.ToString(
                                rs.Fields.Item("ObjType").Value).Trim();

                        config.StageID = Convert.ToInt32(
                            rs.Fields.Item("StageID").Value);

                        config.CurrentStep = Convert.ToInt32(
                            rs.Fields.Item("CurrentStep").Value);

                        return config;
                    }

                    rs.MoveNext();
                }

                // Ada workflow aktif, tetapi tidak ada kondisi
                // template yang cocok dengan objek ini.
                return null;
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }


        //private ItemApprovalConfig GetApprovalConfig(string userCode, string objType)
        //{
        //    SAPbobsCOM.Recordset rs = null;

        //    try
        //    {
        //        rs = (SAPbobsCOM.Recordset)
        //            Program.oCompany.GetBusinessObject(
        //                SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        string safeUser = SqlSafe(userCode);

        //        string safeType = SqlSafe(objType);

        //        string sql = @"
        //    SELECT TOP 1
        //        T.Code AS TemplateCode,
        //        T.Name AS TemplateName,
        //        T.U_TRemarks AS TemplateRemarks,
        //        W.U_U_UserID AS UserCode,
        //        T.U_ObjectType As ObjType,
        //        S.U_U_StageCode AS StageID,
        //        S.U_U_SortId AS CurrentStep
        //    FROM [@WTM1] W
        //    INNER JOIN [@OWTM] T
        //        ON W.Code LIKE T.Code + '-%'
        //    INNER JOIN [@WTM2] S
        //        ON S.Code = W.Code
        //    WHERE W.U_U_UserID = '" + safeUser + @"'
        //      AND T.U_U_Active = 'Y'
        //      AND T.U_ObjectType = '" + safeType + @"'
        //    ORDER BY
        //        S.U_U_SortId ASC,
        //        T.Code DESC";

        //        rs.DoQuery(sql);

        //        if (rs.EoF)
        //        {
        //            throw new Exception(
        //                "Workflow approval tidak ditemukan atau tidak aktif.\n\n" +
        //                "User login : " + userCode + "\n" +
        //                "Object Type: " + objType + "\n\n" +
        //                "Periksa @WTM1, @OWTM, @WTM2 dan status template.");
        //        }

        //        ItemApprovalConfig config = new ItemApprovalConfig();

        //        config.TemplateCode =
        //            Convert.ToString(
        //                rs.Fields.Item("TemplateCode").Value).Trim();

        //        config.TemplateName =
        //            Convert.ToString(
        //                rs.Fields.Item("TemplateName").Value).Trim();

        //        config.TemplateRemarks =
        //            Convert.ToString(
        //                rs.Fields.Item("TemplateRemarks").Value).Trim();

        //        config.UserCode =
        //            Convert.ToString(
        //                rs.Fields.Item("UserCode").Value).Trim();

        //        config.ObjType =
        //            Convert.ToString(
        //                rs.Fields.Item("ObjType").Value).Trim();

        //        config.StageID =
        //            Convert.ToInt32(
        //                rs.Fields.Item("StageID").Value);

        //        config.CurrentStep =
        //            Convert.ToInt32(
        //                rs.Fields.Item("CurrentStep").Value);

        //        return config;
        //    }
        //    finally
        //    {
        //        if (rs != null)
        //            Marshal.ReleaseComObject(rs);
        //    }
        //}

        #endregion

    }
}