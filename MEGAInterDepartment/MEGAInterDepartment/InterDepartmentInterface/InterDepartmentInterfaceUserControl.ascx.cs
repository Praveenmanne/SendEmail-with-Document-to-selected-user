using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Data;
using Microsoft.SharePoint;
using System.Collections;
using System.Net.Mail;
using System.Net;
using System.Net.Mime;
using System.IO;

namespace MEGAInterDepartment.InterDepartmentInterface
{
    public partial class InterDepartmentInterfaceUserControl : UserControl
    {
        #region "Local Variables"
        string toaddress = "";
        #endregion

        #region "Page Events"
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                if (!IsPostBack)
                {
                    BindDocument();
                    BindDepartment();
                }
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , Page_Load()", ex.Message);
            }
        }
        #endregion

        #region "Control Event"
        protected void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                Clear();
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , btnCancel_Click()", ex.Message);
            }
        }
        protected void ddlDepartment_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (ddlDepartment.SelectedItem.Text != "Select")
                {
                    BindEmployee(ddlDepartment.SelectedItem.Text);
                }
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , ddlDepartment_SelectedIndexChanged()", ex.Message);
            }
        }
        protected void btnSave_Click(object sender, EventArgs e)
        {
            if (lbEmployees.Items.Count < 0)
            {
                lblmessage.Text = "Employee must be select";
            }
            else
            {
                lblmessage.Text = "";
                try
                {
                    if (ViewState["dtEmployeeDetail"] != null)
                    {

                        DataTable dtEmployeeDetail = (DataTable)ViewState["dtEmployeeDetail"];
                        for (int i = 0; i < dtEmployeeDetail.Rows.Count; i++)
                        {
                            String Subsite = Convert.ToString(dtEmployeeDetail.Rows[i]["ActualDepartment"]);
                            using (SPSite spSite = new SPSite(SPContext.Current.Web.Site.Url + "/" + Subsite))
                            {
                                using (SPWeb spWeb = spSite.OpenWeb())
                                {

                                    string url = spWeb.Url + "/" + "Shared Documents" + "/" + Convert.ToString(dtEmployeeDetail.Rows[i]["DocumentName"]);
                                    SPFile Document = spWeb.GetFile(url);


                                    if (Document != null)
                                    {
                                        byte[] csvFile = Document.OpenBinary();
                                        string fromadress = SPContext.Current.Web.CurrentUser.Email;
                                        string bodyText = "You have recieved document from " + fromadress;
                                        string subjectText = "Document";
                                        toaddress = GetEmailByEmployeeName(Convert.ToString(dtEmployeeDetail.Rows[i]["EmployeeName"]));

                                        sendMail(fromadress, toaddress, "", bodyText, subjectText, csvFile, Document.Name);

                                    }

                                }
                            }

                        }
                        Clear();
                    }
                }
                catch (Exception ex)
                {
                    MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , btnSave_Click()", ex.Message);
                }
            }

        }
        protected void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                AddEmployeeDetail();
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , btnAdd_Click()", ex.Message);
            }
        }
        protected void gvEmployees_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            DataTable dtEmployeeDetail = new DataTable();
            int rowDeleted = 0;
            try
            {
                if (ViewState["dtEmployeeDetail"] != null)
                {
                    dtEmployeeDetail = (DataTable)ViewState["dtEmployeeDetail"];

                    if (dtEmployeeDetail != null && dtEmployeeDetail.Rows.Count > 0)
                    {
                        gvEmployees.EditIndex = -1;
                        for (int i = 0; i < dtEmployeeDetail.Rows.Count; i++)
                        {
                            if (dtEmployeeDetail.Rows[i].RowState != DataRowState.Deleted)
                            {

                                if (rowDeleted == e.RowIndex)
                                {
                                    dtEmployeeDetail.Rows[i].Delete();
                                }
                                rowDeleted += 1;
                            }
                        }

                        string strSR_NO;
                        int inttotrow = 1;
                        foreach (DataRow drRow in dtEmployeeDetail.Rows)
                        {

                            if (drRow.RowState != DataRowState.Deleted)
                            {
                                strSR_NO = Convert.ToString(inttotrow);
                                strSR_NO = strSR_NO.PadLeft(2, '0');
                                drRow["ID"] = strSR_NO;
                                inttotrow += 1;
                            }
                        }
                        ViewState["dtEmployeeDetail"] = dtEmployeeDetail;
                        gvEmployees.DataSource = dtEmployeeDetail;
                        gvEmployees.DataBind();

                        if (gvEmployees.Rows.Count == 0)
                        {
                            ViewState["dtEmployeeDetail"] = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , gvEmployees_RowDeleting()", ex.Message);
            }
            finally
            {
                if (dtEmployeeDetail != null)
                    dtEmployeeDetail.Dispose();
            }
        }
        protected void gvEmployees_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            try
            {
                if (e.Row.RowType == DataControlRowType.Footer)
                {
                    e.Row.Cells[0].Text = "Page " + (gvEmployees.PageIndex + 1) + " of " + gvEmployees.PageCount;
                }
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , gvEmployees_RowDataBound()", ex.Message);
            }
        }

        #endregion

        #region "Methods"
        //public string GetSubsitefromDepartment(string Department)
        //{
        //    string subsite = "";
        //    DataTable dtGetSubsitefromDepartment = new DataTable();
        //    try
        //    {
        //        using (SPSite spSite = new SPSite(SPContext.Current.Web.Site.Url))
        //        {
        //            using (SPWeb spWeb = spSite.OpenWeb())
        //            {
        //                SPList SPListDepartment = spWeb.Lists["Department"];
        //                SPQuery SPQueryGetSubsitefromDepartment = new SPQuery();

        //                SPQueryGetSubsitefromDepartment.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + Department + "</Value></Eq></Where>";

        //                SPQueryGetSubsitefromDepartment.ViewFields = string.Concat("<FieldRef Name='Title' />",
        //                                                                     "<FieldRef Name='SubSite' />",
        //                                                                     "<FieldRef Name='ID' />");
        //                SPQueryGetSubsitefromDepartment.ViewFieldsOnly = true;

        //                dtGetSubsitefromDepartment = SPListDepartment.GetItems(SPQueryGetSubsitefromDepartment).GetDataTable();
        //                if (dtGetSubsitefromDepartment != null && dtGetSubsitefromDepartment.Rows.Count > 0)
        //                {
        //                    subsite = Convert.ToString(dtGetSubsitefromDepartment.Rows[0]["Subsite"]);
        //                }
        //                return subsite;
        //            }
        //        }
        //    }

        //    catch (Exception ex)
        //    {
        //        MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , GetSubsitefromDepartment()", ex.Message);
        //        return subsite;

        //    }
        //}
        public string GetEmailByEmployeeName(string EmployeeName)
        {
            string fromaddress = "";
            DataTable DtGetEmailByEmployeeName = new DataTable();
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                     {

                         using (SPSite spSite = new SPSite(SPContext.Current.Web.Site.Url))
                         {
                             using (SPWeb spWeb = spSite.OpenWeb())
                             {
                                 SPList SPListMegaEmpCntctList = spWeb.Lists["MegaEmpCntctList"];
                                 SPQuery SPQueryMegaEmpCntctList = new SPQuery();
                                 SPQueryMegaEmpCntctList.Query = "<Where><Eq><FieldRef Name=\"Title\" /><Value Type=\"Text\">" + EmployeeName + "</Value></Eq></Where>";
                                 DtGetEmailByEmployeeName = SPListMegaEmpCntctList.GetItems(SPQueryMegaEmpCntctList).GetDataTable();

                                 if (DtGetEmailByEmployeeName != null && DtGetEmailByEmployeeName.Rows.Count > 0)
                                 {
                                     fromaddress = Convert.ToString(DtGetEmailByEmployeeName.Rows[0]["Email_x0020_Id"]);
                                 }
                             }
                         }
                     });
                return fromaddress;
            }
            catch (Exception ex)
            {
                return fromaddress;
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , GetEmailByEmployeeName()", ex.Message);
            }
            finally
            {
                if (DtGetEmailByEmployeeName != null)
                    DtGetEmailByEmployeeName.Dispose();
            }
        }
        protected void sendMail(string fromaddress, string toAddress, string ccAddress, string bodyText, string subject, byte[] csvFile, string name)
        {
            try
            {
                MemoryStream ms = new MemoryStream(csvFile);
                StreamWriter writer = new StreamWriter(ms);
                writer.Write(csvFile);
                //ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Text.RichText);
                Attachment Docattachment = new System.Net.Mail.Attachment(ms, name);
                Docattachment.ContentDisposition.FileName = name;
                SPSecurity.RunWithElevatedPrivileges(delegate()
                    {
                        using (SPSite spsite = new SPSite(SPContext.Current.Web.Url))
                        {
                            using (SPWeb spWeb = spsite.OpenWeb())
                            {
                                SPListItemCollection mailItmColl = spWeb.Lists["MailConfigurations"].GetItems();
                                if (mailItmColl.Count > 0)
                                {
                                    SPListItem itm = mailItmColl[0];

                                    string strbody = bodyText;

                                    MailMessage mail = new MailMessage();
                                    mail.Body = strbody;
                                    //mail.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8");

                                    mail.From = new MailAddress(fromaddress);
                                    mail.To.Add(new MailAddress(toAddress));

                                    if (ccAddress.Length > 0)
                                    {
                                        string[] ccAddresses = ccAddress.Split(',');
                                        for (int i = 0; i < ccAddresses.Length; i++)
                                        {
                                            mail.CC.Add(new MailAddress(ccAddresses[i]));
                                        }
                                    }
                                    mail.Subject = subject;
                                    mail.IsBodyHtml = true;



                                    mail.Attachments.Add(Docattachment);

                                    SmtpClient client = new SmtpClient(itm["Title"].ToString());

                                    if (Convert.ToString(itm["Username"]).Length > 0 && Convert.ToString(itm["Password"]).Length > 0)
                                    {
                                        string userName = Convert.ToString(itm["Username"]);
                                        string pwd = Convert.ToString(itm["Password"]);
                                        client.Credentials = new NetworkCredential(userName, pwd);
                                        client.Send(mail);
                                    }
                                    else
                                    {
                                        //client.UseDefaultCredentials = true;
                                        client.Send(mail);
                                    }
                                }
                            }
                        }
                    });
                writer.Flush();
                writer.Dispose();
                ms.Close();

            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , sendMail()", ex.Message);
            }
        }
        public void MaintainErroLogs(String ListName, String Description)
        {
            try
            {

                using (SPSite spsite = new SPSite(SPContext.Current.Web.Url))
                {
                    using (SPWeb spWeb = spsite.OpenWeb())
                    {
                        bool AllowUnsafeUpdates = spWeb.AllowUnsafeUpdates;
                        spWeb.AllowUnsafeUpdates = true;
                        SPList SPListErrorLogs = spWeb.Lists["Error Logs"];
                        SPListItem SPListItemErrorLogs = SPListErrorLogs.Items.Add();
                        SPListItemErrorLogs["List Name"] = ListName;
                        SPListItemErrorLogs["Description"] = Description;
                        SPListItemErrorLogs.Update();
                        spWeb.AllowUnsafeUpdates = AllowUnsafeUpdates;
                    }
                }

            }
            catch
            {

            }
        }
        public void BindDepartment()
        {
            ArrayList arrydepartment = new ArrayList();
            DataTable DtGetDepartment = new DataTable();
            DataTable DtAllDepartment = new DataTable();
            DtAllDepartment.Columns.Add("Department");


            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                   {
                       using (SPSite spSite = new SPSite(SPContext.Current.Web.Site.Url))
                       {
                           using (SPWeb spWeb = spSite.OpenWeb())
                           {
                               SPList SPListDepartment = spWeb.Lists["MegaEmpCntctList"];
                               SPQuery SPQueryDepartment = new SPQuery();
                               SPQueryDepartment.Query = "<OrderBy><FieldRef Name='Department' Ascending='True' /></OrderBy>";
                               DtGetDepartment = SPListDepartment.GetItems(SPQueryDepartment).GetDataTable();


                               if (DtGetDepartment != null && DtGetDepartment.Rows.Count > 0)
                               {
                                   for (int i = 0; i < DtGetDepartment.Rows.Count; i++)
                                   {
                                       if (!string.IsNullOrEmpty(Convert.ToString(DtGetDepartment.Rows[i]["Department"])))
                                       {
                                           if (!arrydepartment.Contains(Convert.ToString(DtGetDepartment.Rows[i]["Department"])))
                                           {
                                               arrydepartment.Add(Convert.ToString(DtGetDepartment.Rows[i]["Department"]));
                                           }
                                       }
                                   }
                                   if (arrydepartment.Count > 0)
                                   {
                                       for (int j = 0; j < arrydepartment.Count; j++)
                                       {
                                           DataRow dr = DtAllDepartment.NewRow();
                                           dr["Department"] = arrydepartment[j];
                                           DtAllDepartment.Rows.Add(dr);
                                       }
                                       ddlDepartment.DataSource = DtAllDepartment;
                                       ddlDepartment.DataTextField = "Department";
                                       ddlDepartment.DataBind();
                                       ddlDepartment.Items.Insert(0, "Select");
                                   }
                               }
                               else
                               {
                                   ddlDepartment.Items.Clear();
                                   ddlDepartment.Items.Insert(0, "Select");
                               }
                           }
                       }
                   });
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , BindDepartment()", ex.Message);
            }
            finally
            {
                if (DtGetDepartment != null)
                    DtGetDepartment.Dispose();
            }
        }
        public void BindEmployee(string Department)
        {
            DataTable DtEmployee = new DataTable();
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                      {
                          using (SPSite spSite = new SPSite(SPContext.Current.Web.Site.Url))
                          {
                              using (SPWeb spWeb = spSite.OpenWeb())
                              {
                                  SPList SPListMegaEmpCntctList = spWeb.Lists["MegaEmpCntctList"];
                                  SPQuery SPQueryGetMegaEmpCntctList = new SPQuery();
                                  SPQueryGetMegaEmpCntctList.Query = "<Where><Eq><FieldRef Name='Department' /><Value Type='Lookup'>" + Department + "</Value></Eq></Where>";

                                  DtEmployee = SPListMegaEmpCntctList.GetItems(SPQueryGetMegaEmpCntctList).GetDataTable();

                                  if (DtEmployee != null && DtEmployee.Rows.Count > 0)
                                  {
                                      lbEmployees.DataSource = DtEmployee;
                                      lbEmployees.DataTextField = "Title";
                                      lbEmployees.DataValueField = "ID";
                                      lbEmployees.DataBind();
                                  }
                                  else
                                  {
                                      ddlDepartment.Items.Clear();
                                  }
                              }
                          }
                      });
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , BindEmployee()", ex.Message);
            }
            finally
            {
                if (DtEmployee != null)
                    DtEmployee.Dispose();
            }
        }
        public void BindDocument()
        {
            DataTable DtDocuments = new DataTable();
            DataTable dtAllDocuments = new DataTable();
            dtAllDocuments.Columns.Add("Department");
            dtAllDocuments.Columns.Add("LinkFilename");
            dtAllDocuments.Columns.Add("ID");
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                       {
                           SPGroupCollection grps = SPContext.Current.Web.CurrentUser.Groups;
                           foreach (SPGroup grp in grps)
                           {
                               string groupname = grp.Name.Split(' ')[0];
                               using (SPSite spSite = new SPSite(SPContext.Current.Web.Site.Url + "/" + groupname))
                               {
                                   using (SPWeb spWeb = spSite.OpenWeb())
                                   {
                                       SPList SPListMegaEmpCntctList = spWeb.Lists["Shared Documents"];

                                       SPQuery SPQueryGetDocuments = new SPQuery();
                                       SPQueryGetDocuments.Query = "<Where><Eq><FieldRef Name=\"Status\" /><Value Type=\"Choice\">Approved</Value></Eq></Where><OrderBy><FieldRef Name=\"Title\" Ascending=\"True\" /></OrderBy>";

                                       DtDocuments = SPListMegaEmpCntctList.GetItems(SPQueryGetDocuments).GetDataTable();

                                       if (DtDocuments != null && DtDocuments.Rows.Count > 0)
                                       {
                                           for (int i = 0; i < DtDocuments.Rows.Count; i++)
                                           {
                                               DataRow row = dtAllDocuments.NewRow();
                                               row["LinkFilename"] = Convert.ToString(DtDocuments.Rows[i]["LinkFilename"]);
                                               row["Department"] = groupname;
                                               row["ID"] = Convert.ToString(DtDocuments.Rows[i]["ID"]);
                                               dtAllDocuments.Rows.Add(row);
                                           }
                                       }

                                   }
                               }

                           }
                           if (dtAllDocuments != null && dtAllDocuments.Rows.Count > 0)
                           {
                               ViewState["dtAllDocuments"] = dtAllDocuments;
                               ddlDocuments.DataSource = dtAllDocuments;
                               ddlDocuments.DataTextField = "LinkFilename";
                               ddlDocuments.DataValueField = "ID";
                               ddlDocuments.DataBind();
                               ddlDocuments.Items.Insert(0, "Select");
                           }
                           else
                           {
                               ddlDocuments.Items.Clear();
                               ddlDocuments.Items.Insert(0, "Select");
                           }

                       });
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , BindDocument()", ex.Message);
            }
            finally
            {
                if (DtDocuments != null)
                    DtDocuments.Dispose();
            }

        }
        public void AddEmployeeDetail()
        {
            int TotalSelected = 0;
            foreach (ListItem li in lbEmployees.Items)
            {
                if (li.Selected == true)
                {
                    TotalSelected = TotalSelected + 1;
                }
            }
            if (TotalSelected == 0)
            {
                lblmessage.Text = "Employee must be select";
            }
            else
            {
                lblmessage.Text = "";
            }
            if (TotalSelected > 0)
            {
                DataTable dtEmployeeDetail = new DataTable();
                DataTable dtAllDocuments = (DataTable)ViewState["dtAllDocuments"];
                try
                {
                    int MaxNo;
                    int Count = 0;
                    if (ViewState["dtEmployeeDetail"] != null)
                    {
                        dtEmployeeDetail = (DataTable)ViewState["dtEmployeeDetail"];
                        MaxNo = dtEmployeeDetail.Rows.Count + 1;
                    }
                    else
                    {
                        MaxNo = 1;
                    }
                    if (!dtEmployeeDetail.Columns.Contains("Department"))
                    {
                        dtEmployeeDetail.Columns.Add("ActualDepartment");
                        dtEmployeeDetail.Columns.Add("Department");
                        dtEmployeeDetail.Columns.Add("EmployeeName");
                        dtEmployeeDetail.Columns.Add("DocumentName");
                        dtEmployeeDetail.Columns.Add("DocumentID");
                        dtEmployeeDetail.Columns.Add("ID");

                    }
                    DataView dv = new DataView(dtAllDocuments);
                    dv.RowFilter = dtAllDocuments.DefaultView.RowFilter = "LinkFilename='" + ddlDocuments.SelectedItem.Text + "' AND " +
                                                                                                                                      "ID='" + ddlDocuments.SelectedItem.Value + "'";

                    for (int i = 0; i < dtEmployeeDetail.Rows.Count; i++)
                    {
                        foreach (ListItem li in lbEmployees.Items)
                        {
                            if (li.Selected == true)
                            {
                                if (Convert.ToString(dtEmployeeDetail.Rows[i]["EmployeeName"]) == li.Text && Convert.ToString(dtEmployeeDetail.Rows[i]["Department"]) == ddlDepartment.SelectedItem.Text && ddlDocuments.SelectedItem.Text == Convert.ToString(dtEmployeeDetail.Rows[i]["DocumentName"]))
                                {
                                    lblAddEmployee.Text = "Employee already added.";
                                    Count = 1 + Count;
                                    li.Selected = false;
                                }
                            }
                        }

                    }
                    if (Count == 0)
                    {
                        lblAddEmployee.Text = "";
                        if (ddlDepartment.SelectedItem.Text != "Select")
                        {
                            foreach (ListItem li in lbEmployees.Items)
                            {
                                if (li.Selected == true)
                                {
                                    DataRow dr = dtEmployeeDetail.NewRow();
                                    dr["Department"] = ddlDepartment.SelectedItem.Text;
                                    dr["ID"] = MaxNo;
                                    dr["EmployeeName"] = li.Text;
                                    dr["DocumentName"] = ddlDocuments.SelectedItem.Text;
                                    dr["DocumentID"] = ddlDocuments.SelectedValue;

                                    if (dv.Count > 0 && dv != null)
                                    {
                                        dr["ActualDepartment"] = Convert.ToString(dv.ToTable().Rows[0]["Department"]);
                                    }
                                    dtEmployeeDetail.Rows.Add(dr);
                                    li.Selected = false;
                                }
                            }
                        }
                        ViewState["dtEmployeeDetail"] = dtEmployeeDetail;
                        gvEmployees.DataSource = dtEmployeeDetail;
                        gvEmployees.DataBind();
                        lblAddEmployee.Text = "";
                    }
                }
                catch (Exception ex)
                {
                    MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , AddEmployeeDetail()", ex.Message);
                }
                finally
                {
                    if (dtEmployeeDetail != null)
                        dtEmployeeDetail.Dispose();
                }
            }
        }
        public void Clear()
        {
            try
            {
                ViewState["dtEmployeeDetail"] = null;
                BindDocument();
                BindDepartment();
                gvEmployees.DataSource = null;
                gvEmployees.DataBind();
                lbEmployees.Items.Clear();
            }
            catch (Exception ex)
            {
                MaintainErroLogs("MEGAInterDepartment - InterDepartmentInterface , Clear()", ex.Message);
            }
        }
        #endregion

    }
}
