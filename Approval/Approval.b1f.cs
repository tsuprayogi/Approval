using System;
using System.Globalization;
using System.Runtime.InteropServices;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace Approval
{
    [FormAttribute("Approval.Approval", "Approval.b1f")]
    class Approval : UserFormBase
    {
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;

        private SAPbouiCOM.Matrix matrix0;
        private SAPbouiCOM.Matrix matrix1;

        private SAPbouiCOM.EditText EditApprovalNo;
        private SAPbouiCOM.EditText EditStatus;
        private SAPbouiCOM.EditText EditItemCode;
        private SAPbouiCOM.EditText EditItemName;
        private SAPbouiCOM.EditText EditRequestBy;
        private SAPbouiCOM.EditText EditRequestDate;
        private SAPbouiCOM.EditText EditRemarks;

        private string _currentUserCode = "";
        private string _selectedApprovalCode = "";
        private bool _isLoading = false;

        private SAPbouiCOM.DataTable _dtApprover;

        private SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();


        public Approval()
        {
        }

        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_6").Specific));
            this.matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_11").Specific));
            this.EditApprovalNo = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.EditStatus = ((SAPbouiCOM.EditText)(this.GetItem("Item_10").Specific));
            this.EditItemCode = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.EditItemName = ((SAPbouiCOM.EditText)(this.GetItem("Item_7").Specific));
            this.EditRequestBy = ((SAPbouiCOM.EditText)(this.GetItem("Item_8").Specific));
            this.EditRequestDate = ((SAPbouiCOM.EditText)(this.GetItem("Item_9").Specific));
            this.EditRemarks = ((SAPbouiCOM.EditText)(this.GetItem("Item_12").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {
            try
            {
                SAPbouiCOM.IForm form = UIAPIRawForm;

                if (form == null)
                    throw new Exception("UIAPIRawForm belum tersedia.");

                // Navigasi SAP
                form.EnableMenu("1281", true); // Find
                form.EnableMenu("1282", false); // Add

                SAPbouiCOM.DataTable dt =
                UIAPIRawForm.DataSources.DataTables.Add("DT_APPROVAL");

                dt.Columns.Add("Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10);
                dt.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 120);
                dt.Columns.Add("RequestBy", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("RequestDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 30);
                dt.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 30);
                dt.Columns.Add("Stage", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 30);
                dt.Columns.Add("Step", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10);
                dt.Columns.Add("DocType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);

                matrix0.Columns.Item("colApp")
                .DataBind.Bind("DT_APPROVAL", "Code");

                matrix0.Columns.Item("colType")
                  .DataBind.Bind("DT_APPROVAL", "DocType");

                matrix0.Columns.Item("colItCode")
                    .DataBind.Bind("DT_APPROVAL", "ItemCode");

                matrix0.Columns.Item("colItName")
                    .DataBind.Bind("DT_APPROVAL", "ItemName");

                matrix0.Columns.Item("colReq")
                    .DataBind.Bind("DT_APPROVAL", "RequestBy");

                matrix0.Columns.Item("colDate")
                    .DataBind.Bind("DT_APPROVAL", "RequestDate");

                matrix0.Columns.Item("colStat")
                    .DataBind.Bind("DT_APPROVAL", "Status");

                matrix0.Columns.Item("colRemark")
                   .DataBind.Bind("DT_APPROVAL", "Stage");

                matrix0.Columns.Item("colStep")
                   .DataBind.Bind("DT_APPROVAL", "Step");


                EnsureApprovalDataTables();


                _currentUserCode = GetCurrentUserCode();
                
                ClearDetail();

                LoadApprovalList();

                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "Gagal membuka Approval Form:\n" + ex.Message);
            }
        }

        // =========================================================
        // CURRENT USER
        // =========================================================

        private string GetCurrentUserCode()
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                // 1. Ambil objek DI Company yang valid dari UI API
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                // 2. Sekarang GetBusinessObject() bisa dipanggil dari oCompany
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // 3. Ambil UserSignature dari oCompany DI API
                int userSignature = oCompany.UserSignature;


                string sql =
                    "SELECT USER_CODE " +
                    "FROM OUSR " +
                    "WHERE USERID = " + userSignature;

                rs.DoQuery(sql);
                
                if (rs.RecordCount > 0)
                {
                    return Convert.ToString(
                        rs.Fields.Item("USER_CODE").Value).Trim();
                    
                }

               

                throw new Exception("User SAP tidak ditemukan.");
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        

        private void LoadApprovalList()
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                _isLoading = true;

                matrix0.Clear();
                matrix1.Clear();
                ClearDetail();

                rs = (SAPbobsCOM.Recordset)
                    oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string userCode = SqlSafe(_currentUserCode);

                string sql = @"
            SELECT
                H.Code,
                M.U_T2_ObjName As U_DocType,
                H.U_ItemCode,
                H.U_ItemName,
                H.U_RequestBy,
                H.U_RequestDate,
                H.U_Status,
                K.Remarks,
                H.U_CurrStep
            FROM [@AITM_APV] H
            INNER JOIN [@AITM_APV_L] L            
                ON H.Code = L.Code
            LEFT JOIN OWST K
            ON H.U_StageID = K.WstCode
            LEFT JOIN [@SBO_OBJECT] M ON H.U_DocType = M.U_T2_ObjCode
            WHERE
                H.U_Status = 'PENDING'
                AND L.U_Status = 'PENDING'
                AND L.U_Approver = '" + userCode + @"'
            ORDER BY H.U_RequestDate, H.Code";

                rs.DoQuery(sql);

                SAPbouiCOM.DataTable dt =
                    UIAPIRawForm.DataSources.DataTables.Item("DT_APPROVAL");

                dt.Rows.Clear();

                int row = 0;

                while (!rs.EoF)
                {
                    dt.Rows.Add();

                    dt.SetValue(
                        "Code",
                        row,
                        GetField(rs, "Code"));

                    dt.SetValue(
                        "DocType",
                        row,
                        GetField(rs, "U_DocType"));

                    dt.SetValue(
                        "ItemCode",
                        row,
                        GetField(rs, "U_ItemCode"));

                    dt.SetValue(
                        "ItemName",
                        row,
                        GetField(rs, "U_ItemName"));

                    dt.SetValue(
                        "RequestBy",
                        row,
                        GetField(rs, "U_RequestBy"));

                    dt.SetValue(
                        "RequestDate",
                        row,
                        FormatDate(
                            GetField(rs, "U_RequestDate")));

                    dt.SetValue(
                        "Status",
                        row,
                        GetField(rs, "U_Status"));

                    dt.SetValue(
                       "Stage",
                       row,
                       GetField(rs, "Remarks"));

                    dt.SetValue(
                       "Step",
                       row,
                       GetField(rs, "U_CurrStep"));

                    row++;

                    rs.MoveNext();
                }

                matrix0.LoadFromDataSource();
                matrix0.AutoResizeColumns();
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);

                _isLoading = false;
            }
        }

        // =========================================================
        // LOAD SELECTED APPROVAL DETAIL
        // =========================================================

        private void LoadApprovalDetail(string approvalCode)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                _selectedApprovalCode = approvalCode;

                ClearDetail();

                rs = (SAPbobsCOM.Recordset)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = @"
                    SELECT
                        Code,
                        U_DocType,
                        U_ItemCode,
                        U_ItemName,
                        U_RequestBy,
                        U_RequestDate,
                        U_Status
                    FROM [@AITM_APV]
                    WHERE Code = '" + SqlSafe(approvalCode) + "'";

                rs.DoQuery(sql);

                if (rs.RecordCount == 0)
                    return;

                EditApprovalNo.Value = GetField(rs, "Code");
                EditStatus.Value = GetField(rs, "U_Status");
                EditItemCode.Value = GetField(rs, "U_ItemCode");
                EditItemName.Value = GetField(rs, "U_ItemName");
                EditRequestBy.Value = GetField(rs, "U_RequestBy");
                EditRequestDate.Value =
                    FormatDate(GetField(rs, "U_RequestDate"));

                LoadApproverList(approvalCode);
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        // =========================================================
        // LOAD APPROVER LIST
        // =========================================================

        private void LoadApproverList(string approvalCode)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                matrix1.Clear();

                rs = (SAPbobsCOM.Recordset)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = @"
                    SELECT
                        U_StageID,
                        U_Approver,
                        U_Status,
                        U_Remarks
                    FROM [@AITM_APV_L]
                    WHERE Code = '" + SqlSafe(approvalCode) + @"'
                    ORDER BY U_StageID";

                rs.DoQuery(sql);

                int row = 0;

                while (!rs.EoF)
                {
                    row++;

                    matrix1.AddRow();

                    SetMatrixValue(matrix1, "colApvNo", row,
                        GetField(rs, "U_StageID"));

                    SetMatrixValue(matrix1, "colApprv", row,
                        GetField(rs, "U_Approver"));

                    SetMatrixValue(matrix1, "colApvSt", row,
                        GetField(rs, "U_Status"));

                    SetMatrixValue(matrix1, "colRema", row,
                        GetField(rs, "U_Remarks"));

                    rs.MoveNext();
                }

                matrix1.AutoResizeColumns();
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        // =========================================================
        // EVENT HANDLER
        // =========================================================

        private void SBO_Application_ItemEvent(
            string FormUID,
            ref SAPbouiCOM.ItemEvent pVal,
            out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (FormUID != UIAPIRawForm.UniqueID)
                    return;

                if (pVal.BeforeAction)
                    return;

                if (_isLoading)
                    return;

                // Pilih approval dari matrix utama
                if (pVal.EventType == BoEventTypes.et_CLICK &&
                    pVal.ItemUID == "Item_6" &&
                    pVal.Row > 0)
                {
                    string approvalCode =
                        GetMatrixValue(matrix0, "colApp", pVal.Row);

                    if (!string.IsNullOrWhiteSpace(approvalCode))
                        LoadApprovalDetail(approvalCode);
                }

                //if (pVal.ItemUID == "Item_6" && pVal.ColUID == "colItCode"  && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction == true)
                //{
                //    Application.SBO_Application.MessageBox($"[SUKSES] Event MATRIX_LINK_PRESSED terdeteksi pada Baris: {pVal.Row}, Kolom: {pVal.ColUID}");

                //    // Kita set false dulu agar SAP tidak mencoba membuka form default dan memicu error selama tes
                //    //BubbleEvent = false;
                //    return;
                //}

                //// TEST CASE 2: Jika Test Case 1 tidak muncul, kemungkinan SAP membaca klik panah sebagai ITEM_PRESSED biasa
                //if (pVal.ItemUID == "Item_6" && pVal.ColUID == "colItCode" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == true)
                //{
                //    // Catatan: Beberapa versi SDK membaca klik objek spesifik di dalam matrix sebagai et_ITEM_PRESSED
                //    Application.SBO_Application.MessageBox($"[SUKSES] Event ITEM_PRESSED terdeteksi pada Kolom Linked Button. Baris: {pVal.Row}");

                //    //BubbleEvent = false;
                //    return;
                //}



                // Klik Linked Button pada kolom Item Code
                if (pVal.EventType == BoEventTypes.et_CLICK &&
                    pVal.ItemUID == "Item_6" &&
                    pVal.ColUID == "colItCode" &&
                    pVal.Row > 0)
                {
                    string approvalCode = GetMatrixValue(
                        matrix0,
                        "colApp",
                        pVal.Row);

                    string docCode = GetMatrixValue(
                        matrix0,
                        "colItCode",
                        pVal.Row);

                    if (!string.IsNullOrWhiteSpace(approvalCode) &&
                        !string.IsNullOrWhiteSpace(docCode))
                    {
                        string objType = GetDocTypeFromApproval(approvalCode);

                        if (!string.IsNullOrWhiteSpace(objType))
                        {
                            OpenLinkedObject(objType, docCode);
                        }
                    }

                    return;
                }



                // APPROVE
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                    pVal.ItemUID == "Item_0")
                {
                    ApproveSelected();
                }

                // REJECT
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                    pVal.ItemUID == "Item_1")
                {
                    RejectSelected();
                }

                // REFRESH
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                    pVal.ItemUID == "Item_2")
                {
                    LoadApprovalList();
                }

                // CLOSE
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                    pVal.ItemUID == "Item_3")
                {
                    UIAPIRawForm.Close();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "Error Approval:\n" + ex.Message);
            }
        }

        // =========================================================
        // APPROVE
        // =========================================================

        private void ApproveSelected()
        {
           

            if (string.IsNullOrWhiteSpace(_selectedApprovalCode))
            {
                Application.SBO_Application.MessageBox(
                    "Pilih transaksi approval terlebih dahulu.");
                return;
            }

            string remarks = EditRemarks.Value.Trim();

            SAPbobsCOM.Recordset rs = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string code = SqlSafe(_selectedApprovalCode);
                string userCode = SqlSafe(_currentUserCode);

                // Pastikan user memang approver yang masih pending
                string checkSql = @"
                    SELECT COUNT(*) AS Cnt
                    FROM [@AITM_APV_L]
                    WHERE
                        Code = '" + code + @"'
                        AND U_Approver = '" + userCode + @"'
                        AND U_Status = 'PENDING'";

                rs.DoQuery(checkSql);

                int pendingCount = Convert.ToInt32(
                    rs.Fields.Item("Cnt").Value);

                if (pendingCount == 0)
                {
                    Application.SBO_Application.MessageBox(
                        "Anda tidak memiliki approval pending untuk transaksi ini.");

                    return;
                }

                string updateSql = @"
                    UPDATE [@AITM_APV_L]
                    SET
                        U_Status = 'APPROVED',
                        U_ApprovedBy = '" + userCode + @"',
                        U_ApprovedDate = GETDATE(),                        
                        U_Remarks = '" + SqlSafe(remarks) + @"'
                    WHERE
                        Code = '" + code + @"'
                        AND U_Approver = '" + userCode + @"'
                        AND U_Status = 'PENDING'";

                rs.DoQuery(updateSql);
               

                // Hitung approval yang sudah disetujui
                rs.DoQuery(@"
                    SELECT COUNT(*) AS Cnt
                    FROM [@AITM_APV_L]
                    WHERE
                        Code = '" + code + @"'
                        AND U_Status = 'APPROVED'");


                int approvedCount = Convert.ToInt32(
                    rs.Fields.Item("Cnt").Value);



                // Ambil StageID dari header approval @AITM_APV
                rs.DoQuery(@"
                SELECT U_StageID, U_DocType
                FROM [@AITM_APV]
                WHERE Code = '" + code + @"'");

                if (rs.RecordCount == 0)
                    throw new Exception(
                        "Data approval tidak ditemukan untuk Code: " + code);

                string stageID = Convert.ToString(
                    rs.Fields.Item("U_StageID").Value).Trim();

                string objType = Convert.ToString(
                    rs.Fields.Item("U_DocType").Value).Trim();

                if (string.IsNullOrWhiteSpace(stageID))
                    throw new Exception(
                        "U_StageID belum diisi pada transaksi approval.");

                // Ambil batas approval berdasarkan StageID
                rs.DoQuery(@"
                SELECT MaxReqr
                FROM [OWST]
                WHERE WstCode = '" + SqlSafe(stageID) + @"'");

                if (rs.RecordCount == 0)
                    throw new Exception(
                        "Konfigurasi approval WstCode = " + stageID +
                        " tidak ditemukan.");

                int maxReq = Convert.ToInt32(
                    rs.Fields.Item("MaxReqr").Value);

                if (approvedCount >= maxReq)
                {
                    rs.DoQuery(@"
                        UPDATE [@AITM_APV]
                        SET U_Status = 'APPROVED', U_StageStatus = 'APPROVED'
                        WHERE Code = '" + code + @"'");

                    string itemCode = EditItemCode.Value.Trim();

                    // Aktivasi berdasarkan Object Type
                    if (objType == "4")
                    {
                        ActivateItem(itemCode);
                    }
                    else if (objType == "2")
                    {
                        ActivateBPCode(itemCode);
                    }
                    else
                    {
                        throw new Exception(
                            "Object Type tidak didukung: " + objType);
                    }



                    Application.SBO_Application.MessageBox(
                        "Approval selesai. Item telah diaktifkan.");
                }
                else
                {
                    Application.SBO_Application.MessageBox(
                        "Approval berhasil disimpan. Menunggu approver lainnya.");
                }

                LoadApprovalList();
                ClearDetail();
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        // =========================================================
        // REJECT
        // =========================================================

        private void RejectSelected()
        {
            if (string.IsNullOrWhiteSpace(_selectedApprovalCode))
            {
                Application.SBO_Application.MessageBox(
                    "Pilih transaksi approval terlebih dahulu.");
                return;
            }

            string remarks = EditRemarks.Value.Trim();

            SAPbobsCOM.Recordset rs = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string code = SqlSafe(_selectedApprovalCode);
                string userCode = SqlSafe(_currentUserCode);

                rs.DoQuery(@"
                    SELECT COUNT(*) AS Cnt
                    FROM [@AITM_APV_L]
                    WHERE
                        Code = '" + code + @"'
                        AND U_Approver = '" + userCode + @"'
                        AND U_Status = 'PENDING'");

                int pendingCount = Convert.ToInt32(
                    rs.Fields.Item("Cnt").Value);

                if (pendingCount == 0)
                {
                    Application.SBO_Application.MessageBox(
                        "Anda tidak memiliki approval pending untuk transaksi ini.");

                    return;
                }

                rs.DoQuery(@"
                    UPDATE [@AITM_APV_L]
                    SET
                        U_Status = 'REJECTED',
                        U_ApprovedBy = '" + userCode + @"',
                        U_ApprovedDate = GETDATE(),
                        U_Remarks = '" + SqlSafe(remarks) + @"'
                    WHERE
                        Code = '" + code + @"'
                        AND U_Approver = '" + userCode + @"'
                        AND U_Status = 'PENDING'");

                //rs.DoQuery(@"
                //    UPDATE [@AITM_APV]
                //    SET U_Status = 'REJECTED'
                //    WHERE Code = '" + code + @"'");

                rs.DoQuery(@"
                    UPDATE [@AITM_APV]
                    SET U_Status = 'REJECTED', U_StageStatus = 'REJECTED'
                    WHERE Code = '" + code + @"'");

                // Ambil ItemCode dari detail approval
                rs.DoQuery(@"
                    SELECT U_ItemCode, U_DocType
                    FROM [@AITM_APV]
                    WHERE Code = '" + code + @"'");

                if (rs.RecordCount == 0)
                    throw new Exception(
                        "Data approval tidak ditemukan: " + code);

                string itemCode = Convert.ToString(
                    rs.Fields.Item("U_ItemCode").Value).Trim();

                string objType = Convert.ToString(
                    rs.Fields.Item("U_DocType").Value).Trim();

                // Update status approval di Item Master menjadi R
                RejectItem(itemCode, objType);

                Application.SBO_Application.MessageBox(
                   "Transaksi approval ditolak. " +
                    "Status Item Master telah diubah menjadi Rejected.");


                LoadApprovalList();
                ClearDetail();
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }



        private void RejectItem(string itemCode, string objType)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(itemCode))
                    throw new Exception("Kode item atau BP tidak boleh kosong.");

                // Object Type 4 = Item Master Data
                if (objType == "4")
                {
                    SAPbobsCOM.Items item = null;

                    try
                    {
                        item = (SAPbobsCOM.Items)
                            Program.oCompany.GetBusinessObject(
                                SAPbobsCOM.BoObjectTypes.oItems);

                        if (!item.GetByKey(itemCode))
                            throw new Exception(
                                "Item tidak ditemukan: " + itemCode);

                        // Item tetap tidak aktif setelah ditolak
                        item.Valid = SAPbobsCOM.BoYesNoEnum.tNO;
                        item.Frozen = SAPbobsCOM.BoYesNoEnum.tYES;

                        // R = Rejected
                        item.UserFields.Fields
                            .Item("U_ApprovalStatus").Value = "R";

                        int result = item.Update();

                        if (result != 0)
                        {
                            int errorCode;
                            string errorMessage;

                            Program.oCompany.GetLastError(
                                out errorCode,
                                out errorMessage);

                            throw new Exception(
                                "Gagal menolak Item: " +
                                errorCode + " - " + errorMessage);
                        }
                    }
                    finally
                    {
                        if (item != null)
                        {
                            System.Runtime.InteropServices.Marshal
                                .ReleaseComObject(item);
                        }
                    }
                }
                // Object Type 2 = Business Partner
                else if (objType == "2")
                {
                    SAPbobsCOM.BusinessPartners bp = null;

                    try
                    {
                        bp = (SAPbobsCOM.BusinessPartners)
                            Program.oCompany.GetBusinessObject(
                                SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                        if (!bp.GetByKey(itemCode))
                            throw new Exception(
                                "Business Partner tidak ditemukan: " + itemCode);

                        // Tandai approval sebagai Rejected
                        bp.UserFields.Fields
                            .Item("U_ApprovalStatus").Value = "R";

                        int result = bp.Update();

                        if (result != 0)
                        {
                            int errorCode;
                            string errorMessage;

                            Program.oCompany.GetLastError(
                                out errorCode,
                                out errorMessage);

                            throw new Exception(
                                "Gagal menolak Business Partner: " +
                                errorCode + " - " + errorMessage);
                        }
                    }
                    finally
                    {
                        if (bp != null)
                        {
                            System.Runtime.InteropServices.Marshal
                                .ReleaseComObject(bp);
                        }
                    }
                }
                else
                {
                    throw new Exception(
                        "Object Type tidak didukung: " + objType);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "Gagal melakukan reject: " + ex.Message);
            }
        }


        private void ActivateItem(string itemCode)
        {
            SAPbobsCOM.Items item = null;

            try
            {
                item = (SAPbobsCOM.Items)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.oItems);

                if (!item.GetByKey(itemCode))
                    throw new Exception(
                        "Item tidak ditemukan: " + itemCode);

                // =====================================================
                // DEBUG
                // =====================================================
                Application.SBO_Application.StatusBar.SetText(
                    "Item : " + item.ItemCode +                    
                    "\nApproval Status : " +
                    Convert.ToString(
                        item.UserFields.Fields
                            .Item("U_ApprovalStatus").Value),
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                // =====================================================
                // AKTIFKAN ITEM
                // =====================================================
                item.Valid = SAPbobsCOM.BoYesNoEnum.tYES;
                item.Frozen = SAPbobsCOM.BoYesNoEnum.tNO;

                // =====================================================
                // UBAH STATUS APPROVAL
                // P = Pending
                // A = Approved
                // =====================================================
                item.UserFields.Fields
                    .Item("U_ApprovalStatus")
                    .Value = "A";

                int result = item.Update();

                if (result != 0)
                {
                    int errorCode;
                    string errorMessage;

                    Program.oCompany.GetLastError(
                        out errorCode,
                        out errorMessage);

                    throw new Exception(
                        "Gagal mengaktifkan item: " +
                        errorCode + " - " + errorMessage);
                }

                // =====================================================
                // DEBUG BERHASIL
                // =====================================================
                Application.SBO_Application.StatusBar.SetText(
                    "ITEM BERHASIL DI-APPROVE\n\n" +
                    "Item : " + itemCode +                   
                    "\nApproval Status : A",
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            finally
            {
                if (item != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
            }
        }

        private void ActivateBPCode(string itemCode)
        {
            SAPbobsCOM.BusinessPartners partners = null;

            try
            {
                partners = (SAPbobsCOM.BusinessPartners)
                    Program.oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                if (!partners.GetByKey(itemCode))
                    throw new Exception(
                        "Item tidak ditemukan: " + itemCode);

                // =====================================================
                // DEBUG
                // =====================================================
                Application.SBO_Application.StatusBar.SetText(
                    "BPCode: " + partners.CardCode +
                    "\nApproval Status : " +
                    Convert.ToString(
                        partners.UserFields.Fields
                            .Item("U_ApprovalStatus").Value),
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                // =====================================================
                // AKTIFKAN ITEM
                // =====================================================
                partners.Valid = SAPbobsCOM.BoYesNoEnum.tYES;
                partners.Frozen = SAPbobsCOM.BoYesNoEnum.tNO;

                // =====================================================
                // UBAH STATUS APPROVAL
                // P = Pending
                // A = Approved
                // =====================================================
                partners.UserFields.Fields
                    .Item("U_ApprovalStatus")
                    .Value = "A";

                int result = partners.Update();

                if (result != 0)
                {
                    int errorCode;
                    string errorMessage;

                    Program.oCompany.GetLastError(
                        out errorCode,
                        out errorMessage);

                    throw new Exception(
                        "Gagal mengaktifkan item: " +
                        errorCode + " - " + errorMessage);
                }

                // =====================================================
                // DEBUG BERHASIL
                // =====================================================
                Application.SBO_Application.StatusBar.SetText(
                    "BP Master BERHASIL DI-APPROVE\n\n" +
                    "BP Code : " + itemCode +
                    "\nApproval Status : A",
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            finally
            {
                if (partners != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(partners);
            }
        }


        // =========================================================
        // CLEAR DETAIL
        // =========================================================

        private void ClearDetail()
        {
            //_selectedApprovalCode = "";

            EditApprovalNo.Value = "";
            EditStatus.Value = "";
            EditItemCode.Value = "";
            EditItemName.Value = "";
            EditRequestBy.Value = "";
            EditRequestDate.Value = "";
            EditRemarks.Value = "";

            matrix1.Clear();
        }

       

        private void SetMatrixValue(
            SAPbouiCOM.Matrix matrix,
            string columnUID,
            int row,
            string value)
        {
            try
            {
                SAPbouiCOM.EditText cell =
                    (SAPbouiCOM.EditText)
                    matrix.Columns.Item(columnUID)
                        .Cells.Item(row)
                        .Specific;

                cell.Value = value ?? "";
            }
            catch (Exception ex)
            {
                throw new Exception(
                    $"Gagal isi Matrix. Column={columnUID}, Row={row}, Value={value}\n{ex.Message}",
                    ex);
            }
        }

        private string GetMatrixValue(
            SAPbouiCOM.Matrix matrix,
            string columnUID,
            int row)
        {
            return ((SAPbouiCOM.EditText)
                matrix.Columns.Item(columnUID)
                    .Cells.Item(row).Specific).Value.Trim();
        }

        private string GetField(
            SAPbobsCOM.Recordset rs,
            string fieldName)
        {
            object value = rs.Fields.Item(fieldName).Value;

            if (value == null || value == DBNull.Value)
                return "";

            return Convert.ToString(value).Trim();
        }

        private string FormatDate(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            DateTime date;

            if (DateTime.TryParse(value, out date))
                return date.ToString("yyyy-MM-dd");

            return value;
        }

        private string SqlSafe(string value)
        {
            return (value ?? "").Replace("'", "''");
        }


        private void EnsureApprovalDataTables()
        {
            SAPbouiCOM.IForm oForm = UIAPIRawForm;

            try
            {
                _dtApprover = oForm.DataSources.DataTables.Item("DT_APPROVER");
            }
            catch
            {
                _dtApprover = oForm.DataSources.DataTables.Add("DT_APPROVER");

                _dtApprover.Columns.Add(
                    "ApprovalNo",
                    SAPbouiCOM.BoFieldsType.ft_AlphaNumeric,
                    50);

                _dtApprover.Columns.Add(
                    "Approver",
                    SAPbouiCOM.BoFieldsType.ft_AlphaNumeric,
                    100);

                _dtApprover.Columns.Add(
                    "Status",
                    SAPbouiCOM.BoFieldsType.ft_AlphaNumeric,
                    50);

                _dtApprover.Columns.Add(
                    "Remarks",
                    SAPbouiCOM.BoFieldsType.ft_AlphaNumeric,
                    254);
            }

            matrix1.Columns.Item("colApvNo")
                .DataBind.Bind("DT_APPROVER", "ApprovalNo");

            matrix1.Columns.Item("colApprv")
                .DataBind.Bind("DT_APPROVER", "Approver");

            matrix1.Columns.Item("colApvSt")
                .DataBind.Bind("DT_APPROVER", "Status");

            matrix1.Columns.Item("colRema")
                .DataBind.Bind("DT_APPROVER", "Remarks");
        }



        private string GetDocTypeFromApproval(string approvalCode)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                rs = (SAPbobsCOM.Recordset)
                    oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string safeCode = SqlSafe(approvalCode);

                string sql = @"
            SELECT U_DocType
            FROM [@AITM_APV]
            WHERE Code = '" + safeCode + "'";

                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    return Convert.ToString(
                        rs.Fields.Item("U_DocType").Value
                    ).Trim();
                }

                return string.Empty;
            }
            finally
            {
                if (rs != null)
                    System.Runtime.InteropServices.Marshal
                        .ReleaseComObject(rs);
            }
        }


        private void OpenLinkedObject(string objType, string docCode)
        {
            try
            {
                switch (objType.Trim())
                {
                    case "4": // Item Master Data
                        Application.SBO_Application.OpenForm(
                            SAPbouiCOM.BoFormObjectEnum.fo_Items,
                            "",
                            docCode);
                        break;

                    case "2": // Item Master Data
                        Application.SBO_Application.OpenForm(
                            SAPbouiCOM.BoFormObjectEnum.fo_BusinessPartner,
                            "",
                            docCode);
                        break;

                    default:
                        Application.SBO_Application.StatusBar.SetText(
                            "ObjType " + objType +
                            " belum didukung.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        break;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Gagal membuka objek: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
    }
}