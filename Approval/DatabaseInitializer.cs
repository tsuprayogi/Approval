using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;
using Application = SAPbouiCOM.Framework.Application;

namespace Approval
{
    public static class DatabaseInitializer
    {
        public static void Init()
        {
            Company company = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

            CreateApprovalHeader(company);
            CreateApprovalLine(company);
            CreateApprovalTemplate(company);
            CreateApprovalTemplateUsers(company);
            CreateApprovalTemplateStages(company);
            CreateSBOObject(company);

            CreateItemMasterUDFs(company);
            CreateBPMasterUDFs(company);

            //CreateUDO(company);
        }

        // =====================================================
        // HEADER : @AITM_APV
        // =====================================================
        static void CreateApprovalHeader(Company company)
        {
            if (!TableExists(company, "@AITM_APV"))
            {
                UserTablesMD udt = null;

                try
                {
                    udt = (UserTablesMD)
                        company.GetBusinessObject(BoObjectTypes.oUserTables);

                    udt.TableName = "AITM_APV";
                    udt.TableDescription = "Approval Header";
                    udt.TableType = BoUTBTableType.bott_NoObject;

                    Check(company, udt.Add());
                }
                finally
                {
                    Release(udt);
                }
            }

            AddField(company, "@AITM_APV", "ItemCode", "Item Code",
                BoFieldTypes.db_Alpha, 50);

            AddField(company, "@AITM_APV", "ItemName", "Item Name",
                BoFieldTypes.db_Alpha, 100);

            AddField(company, "@AITM_APV", "Status", "Approval Status",
                BoFieldTypes.db_Alpha, 20);

            AddField(company, "@AITM_APV", "RequestBy", "Request By",
                BoFieldTypes.db_Alpha, 50);

            AddField(company, "@AITM_APV", "RequestDate", "Request Date",
                BoFieldTypes.db_Date);

            AddField(company, "@AITM_APV", "StageID", "Stage ID",
                BoFieldTypes.db_Alpha, 20);

            AddField(company, "@AITM_APV", "CurrStep", "Current Step",
                BoFieldTypes.db_Numeric);

            AddField(company, "@AITM_APV", "StageStatus", "Stage Status",
                BoFieldTypes.db_Alpha, 20);

            AddField(company, "@AITM_APV", "DocType", "Object Code",
                BoFieldTypes.db_Alpha, 20);
        }



        // =====================================================
        // DETAIL : @AITM_APV_L
        // =====================================================
        static void CreateApprovalLine(Company company)
        {
            if (!TableExists(company, "@AITM_APV_L"))
            {
                UserTablesMD udt = null;

                try
                {
                    udt = (UserTablesMD)
                        company.GetBusinessObject(BoObjectTypes.oUserTables);

                    udt.TableName = "AITM_APV_L";
                    udt.TableDescription = "Approval Detail";
                    udt.TableType = BoUTBTableType.bott_NoObject;

                    Check(company, udt.Add());
                }
                finally
                {
                    Release(udt);
                }
            }


            AddField(company, "@AITM_APV_L", "StageID", "Stage ID",
                BoFieldTypes.db_Alpha, 20);

            AddField(company, "@AITM_APV_L", "Approver", "Approver",
                BoFieldTypes.db_Alpha, 50);

            AddField(company, "@AITM_APV_L", "Status", "Approval Status",
                BoFieldTypes.db_Alpha, 20);

            AddField(company, "@AITM_APV_L", "ApprovedBy", "Approved By",
                BoFieldTypes.db_Alpha, 50);

            AddField(company, "@AITM_APV_L", "ApprovedDate", "Approved Date",
                BoFieldTypes.db_Date);

            AddField(company, "@AITM_APV_L", "Remarks", "Remarks",
                BoFieldTypes.db_Alpha, 254);
          

        }

        //Create BOM Version on OWOR
        static void CreateApprovalTemplate(Company company)
        {
            if (!TableExists(company, "@OWTM"))
            {
                UserTablesMD udt = null;

                try
                {
                    udt = (UserTablesMD)
                        company.GetBusinessObject(
                            BoObjectTypes.oUserTables);

                    udt.TableName = "OWTM";
                    udt.TableDescription = "Approval Template Header";
                    udt.TableType = BoUTBTableType.bott_NoObject;

                    Check(company, udt.Add());
                }
                finally
                {
                    Release(udt);
                }
            }

            AddField(company, "@OWTM", "U_Active", "Active",
                BoFieldTypes.db_Alpha, 1);

            AddField(company, "@OWTM", "TName", "Template Name",
                BoFieldTypes.db_Alpha, 100);

            AddField(company, "@OWTM", "TRemarks", "Template Remarks",
                BoFieldTypes.db_Alpha, 254);

            AddField(company, "@OWTM", "ObjectType", "Object Type",
                BoFieldTypes.db_Alpha, 20);

            AddField(company, "@OWTM", "TCondition", "Condition",
                BoFieldTypes.db_Alpha, 250);
        }


        static void CreateApprovalTemplateUsers(Company company)
        {
            if (!TableExists(company, "@WTM1"))
            {
                UserTablesMD udt = null;

                try
                {
                    udt = (UserTablesMD)
                        company.GetBusinessObject(
                            BoObjectTypes.oUserTables);

                    udt.TableName = "WTM1";
                    udt.TableDescription = "Approval Template Users";
                    udt.TableType = BoUTBTableType.bott_NoObject;

                    Check(company, udt.Add());
                }
                finally
                {
                    Release(udt);
                }
            }

            AddField(company, "@WTM1", "U_UserID", "User ID",
                BoFieldTypes.db_Alpha, 50);

            AddField(company, "@WTM1", "U_Dept", "Department",
                BoFieldTypes.db_Alpha, 20);
        }



        static void CreateSBOObject(Company company)
        {
            if (!TableExists(company, "@SBO_OBJECT"))
            {
                UserTablesMD udt = null;

                try
                {
                    udt = (UserTablesMD)company.GetBusinessObject(
                        BoObjectTypes.oUserTables);

                    udt.TableName = "SBO_OBJECT";
                    udt.TableDescription = "Tabel Object dan Form";
                    udt.TableType = BoUTBTableType.bott_MasterData;

                    Check(company, udt.Add());
                }
                finally
                {
                    Release(udt);
                }
            }

            AddFieldWithValidValues(
                   company,
                   "@SBO_OBJECT",
                   "T2_ObjGroup",
                   "Object Group",
                   BoFieldTypes.db_Alpha,
                   30,
                   new[]
                   {
                        ("Master Data", "Master Data"),
                        ("Transaction", "Transaction"),
                        ("Custom Object", "Custom Object")
                   });

            // Field konfigurasi Object dan Form
            AddField(company, "@SBO_OBJECT", "T2_ObjCode",
                "Object Code", BoFieldTypes.db_Alpha, 10);

            AddField(company, "@SBO_OBJECT", "T2_ObjBase",
                "Object Base", BoFieldTypes.db_Alpha, 10);

            AddField(company, "@SBO_OBJECT", "T2_ObjName",
                "Object Name", BoFieldTypes.db_Alpha, 50);

            AddField(company, "@SBO_OBJECT", "T2_ObjTable",
                "Object Table", BoFieldTypes.db_Alpha, 10);

            AddField(company, "@SBO_OBJECT", "T2_FormId",
                "Form ID", BoFieldTypes.db_Alpha, 10);

            
        }


        static void CreateApprovalTemplateStages(Company company)
        {
            if (!TableExists(company, "@WTM2"))
            {
                UserTablesMD udt = null;

                try
                {
                    udt = (UserTablesMD)
                        company.GetBusinessObject(
                            BoObjectTypes.oUserTables);

                    udt.TableName = "WTM2";
                    udt.TableDescription = "Approval Template Stages";
                    udt.TableType = BoUTBTableType.bott_NoObject;

                    Check(company, udt.Add());
                }
                finally
                {
                    Release(udt);
                }
            }

            AddField(company, "@WTM2", "U_StageCode", "Stage Code",
                BoFieldTypes.db_Alpha, 50);

            AddField(company, "@WTM2", "U_SortId", "Sort ID",
                BoFieldTypes.db_Numeric);

            AddField(company, "@WTM2", "U_Remarks", "Stage Remarks",
                BoFieldTypes.db_Alpha, 250);

            AddField(company, "@WTM2", "U_Approver", "Stage Remarks",
                BoFieldTypes.db_Alpha, 250);
        }


        static void CreateItemMasterUDFs(Company company)
        {
            

            AddFieldWithValidValues(
                   company,
                   "OITM",
                   "ApprovalStatus",
                   "Approval Status",
                   BoFieldTypes.db_Alpha,
                   20,
                   new[]
                   {                        
                        ("P", "PENDING"),
                        ("A", "APPROVE"),
                        ("R", "REJECT")
                   });


            AddField(company, "OITM", "ApprovalDocEntry",
                     "Approval DocEntry", BoFieldTypes.db_Alpha, 10);

            AddField(company, "OITM", "ApprovalUser",
                     "Approval Created By", BoFieldTypes.db_Alpha, 30);


        }

        static void CreateBPMasterUDFs(Company company)
        {


            AddFieldWithValidValues(
                   company,
                   "OCRD",
                   "ApprovalStatus",
                   "Approval Status",
                   BoFieldTypes.db_Alpha,
                   20,
                   new[]
                   {
                        ("P", "PENDING"),
                        ("A", "APPROVE"),
                        ("R", "REJECT")
                   });


            AddField(company, "OCRD", "ApprovalDocEntry",
                     "Approval DocEntry", BoFieldTypes.db_Alpha, 10);

            AddField(company, "OCRD", "ApprovalUser",
                     "Approval Created By", BoFieldTypes.db_Alpha, 30);


        }
        

        static void CreateUDO(Company company)
        {
            UserObjectsMD udo = null;

            try
            {
                udo = (UserObjectsMD)company.GetBusinessObject(
                    BoObjectTypes.oUserObjectsMD);

                bool exists = udo.GetByKey("DF04");

                if (!exists)
                {
                    udo.Code = "DF04";
                    udo.Name = "SBO Object";
                    udo.TableName = "SBO_OBJECT";
                    udo.ObjectType = BoUDOObjType.boud_MasterData;
                }

                // Fitur UDO
                udo.CanCreateDefaultForm = BoYesNoEnum.tYES;
                udo.CanFind = BoYesNoEnum.tYES;
                udo.CanDelete = BoYesNoEnum.tYES;

                // Konfigurasi menu
                udo.MenuItem = BoYesNoEnum.tYES;
                udo.MenuCaption = "SBO Object";
                udo.MenuUID = "DF04";

                // Menu induk: pastikan ID ini valid pada SAP B1 Anda.
                // Contoh resmi SAP menggunakan 43535 untuk Business Partners.
                // Ganti sesuai lokasi menu yang diinginkan.
               // udo.FatherMenuID = 47616;
                //udo.Position = 0;

                

                // Tidak menggunakan penomoran series otomatis
                // karena Code dapat diatur oleh aplikasi.
                udo.ManageSeries = BoYesNoEnum.tNO;

                int result = exists
                    ? udo.Update()
                    : udo.Add();

                if (result != 0)
                {
                    company.GetLastError(
                        out int errCode,
                        out string errMsg);

                    throw new Exception(
                        $"Gagal {(exists ? "update" : "membuat")} UDO " +
                        $"DF04: {errCode} - {errMsg}");
                }

                Application.SBO_Application.StatusBar.SetText(
                    exists
                        ? "UDO DF04 berhasil diperbarui."
                        : "UDO DF04 berhasil dibuat.",
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                Release(udo);
            }
        }





        // =====================================================
        // CHECK TABLE
        // =====================================================
        static bool TableExists(Company company, string tableName)
        {
            Recordset rs = null;

            try
            {
                rs = (Recordset)
                    company.GetBusinessObject(BoObjectTypes.BoRecordset);

                rs.DoQuery($@"
                    SELECT 1
                    FROM OUTB
                    WHERE TableName = '{tableName.Replace("@", "")}'
                ");

                return !rs.EoF;
            }
            finally
            {
                Release(rs);
            }
        }
        
        static bool FieldExists(Company company, string table, string field)
        {
            Recordset rs = null;

            try
            {
                rs = (Recordset)
                    company.GetBusinessObject(BoObjectTypes.BoRecordset);

                string alias1 = field;
                string alias2 = field.StartsWith("U_")
                    ? field
                    : "U_" + field;

                rs.DoQuery($@"
            SELECT TOP 1 AliasID
            FROM CUFD
            WHERE TableID = '{table}'
            AND AliasID IN ('{alias1}','{alias2}')
        ");

                return !rs.EoF;
            }
            finally
            {
                Release(rs);
            }
        }
        // =====================================================
        // ADD FIELD (SAFE)
        // =====================================================
        static void AddField(Company company, string table, string field,
                             string desc, BoFieldTypes type, int size = 0)
        {
            if (FieldExists(company, table, field))
                return;

            UserFieldsMD udf = null;

            try
            {
                udf = (UserFieldsMD)
                    company.GetBusinessObject(BoObjectTypes.oUserFields);

                udf.TableName = table;
                udf.Name = field;
                udf.Description = desc;
                udf.Type = type;

                if (type == BoFieldTypes.db_Alpha)
                    udf.Size = size;

                Check(company, udf.Add());
            }
            finally
            {
                Release(udf);
            }
        }


        static void AddFieldWithValidValues(
               Company company,
               string table,
               string field,
               string desc,
               BoFieldTypes type,
               int size,
               (string Value, string Desc)[] validValues,
               string defaultValue = null
                )
        {
            if (FieldExists(company, table, field))
                return;

            UserFieldsMD udf = null;

            try
            {
                udf = (UserFieldsMD)
                    company.GetBusinessObject(BoObjectTypes.oUserFields);

                udf.TableName = table;
                udf.Name = field;
                udf.Description = desc;
                udf.Type = type;

                if (type == BoFieldTypes.db_Alpha)
                    udf.Size = size;

                // ===== Add Valid Values =====
                foreach (var vv in validValues)
                {
                    udf.ValidValues.Value = vv.Value;
                    udf.ValidValues.Description = vv.Desc;
                    udf.ValidValues.Add();
                }

                if (!string.IsNullOrEmpty(defaultValue))
                    udf.DefaultValue = defaultValue;

                Check(company, udf.Add());
            }
            finally
            {
                Release(udf);
            }
        }



        // =====================================================
        // ERROR HANDLER
        // =====================================================
        static void Check(Company company, int ret)
        {


            if (ret != 0)
            {
                company.GetLastError(out int errCode, out string errMsg);
                throw new System.Exception($"{errCode} - {errMsg}");
            }
        }

        // =====================================================
        // COM RELEASE HELPER (KUNCI 1120)
        // =====================================================
        static void Release(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
        }
    }
}

