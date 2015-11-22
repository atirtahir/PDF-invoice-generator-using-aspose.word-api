using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using Aspose.Words.Fields;
using Aspose.Drawing;
using System.Data;

namespace AsposeWordApplication
{

    public partial class _Default : Page
    {

      
        
        protected void Page_Load(object sender, EventArgs e)
        { 
        //ExStart:LoadPage

            //new code

            if (!Page.IsPostBack)
            {
                List<int> lst = new List<int>();
                lst.Add(1);
                lst.Add(2);
                lst.Add(3);
                ItemDetails.DataSource = lst;
                ItemDetails.DataBind();
            }
            //ExEnd:LoadPage
        }
        

        protected void Button1_Click(object sender, EventArgs e)
        {

            string templatePath = Server.MapPath("~/Files/");
            //string logoPath = Server.MapPath("~/App_Data/");
            DocumentBuilder builder = new DocumentBuilder();
            Document doc = new Document(templatePath + "MailMerge Template.docx");
             
            //upload company logo
            if (Request.Files.Count > 0)
            {
                var file2 = Request.Files[0];

                if (file2 != null && file2.ContentLength > 0)
                {

                    var fileName = Path.GetFileName(file2.FileName);
                    var pathOfLogo = Path.Combine(Server.MapPath("~/App_Data/"), fileName);
                    file2.SaveAs(pathOfLogo);
                    string logoPath = Server.MapPath("~/App_Data/"); 


                    DataSet dSet = new DataSet();
                    DataTable dTable = new DataTable("Item");
                    //adding columns in table
                    dTable.Columns.Add("Name");
                    dTable.Columns.Add("Quantity");
                    dTable.Columns.Add("Price");
                    //looping datagridview and fetching data from each row
                    for (int i = 0; i < ItemDetails.Rows.Count; i++)
                    {
                        TextBox txtName = ItemDetails.Rows[i].FindControl("txtName") as TextBox;
                        TextBox txtPrice = ItemDetails.Rows[i].FindControl("txtPrice") as TextBox;
                        TextBox txtItem = ItemDetails.Rows[i].FindControl("txtItem") as TextBox;
                        //assigning fetched data from gridview to datarow
                        DataRow dr = dTable.NewRow();
                        dr["Name"] = txtName.Text;
                        dr["Quantity"] = txtItem.Text;
                        dr["Price"] = txtPrice.Text;
                        
                        dTable.Rows.Add(dr);

                    }
                    //binding table data to dataset
                    dSet.Tables.Add(dTable);
                    //merging company's information and client's information
                    doc.MailMerge.Execute(new string[] { "CompanyName", "Email", "Address", "MyImage", "ClientName", "ClientEmail", "ClientAddress" }, 
                        new object[] { comname.Value, compmail.Value, add.Value, File.ReadAllBytes(pathOfLogo), cname.Value, mail.Value, clientaddress.Value});
                    //merging data to mergefield table
                    doc.MailMerge.ExecuteWithRegions(dSet);


                    doc.Save(templatePath + "Output.pdf");
                    string filename = "Output.pdf";
                    Response.Redirect("~/Files/" + filename + "");
                }

                //if user is not attaching logo

                else if (file2.ContentLength <= 0)
                {
                   
                    DataSet dSet = new DataSet();
                    DataTable dTable = new DataTable("Item");
                    //adding columns in table
                    dTable.Columns.Add("Name");
                    dTable.Columns.Add("Quantity");
                    dTable.Columns.Add("Price");
                    //looping datagridview and fetching data from each row
                    for (int i = 0; i < ItemDetails.Rows.Count; i++)
                    {
                        TextBox txtName = ItemDetails.Rows[i].FindControl("txtName") as TextBox;
                        TextBox txtPrice = ItemDetails.Rows[i].FindControl("txtPrice") as TextBox;
                        TextBox txtItem = ItemDetails.Rows[i].FindControl("txtItem") as TextBox;
                     
                        DataRow dr = dTable.NewRow();
                        //assigning fetched data from gridview to datarow
                        dr["Name"] = txtName.Text;
                        dr["Quantity"] = txtItem.Text;
                        dr["Price"] = txtPrice.Text;
                        dTable.Rows.Add(dr);

                    }

                    dSet.Tables.Add(dTable);
                    //merging company's information and client's information
                    doc.MailMerge.Execute(new string[] { "CompanyName", "Email", "Address", "ClientName", "ClientEmail", "ClientAddress" },
                       new object[] { comname.Value, compmail.Value, add.Value, cname.Value, mail.Value, clientaddress.Value });
                    //merging data to mergefield table
                    doc.MailMerge.ExecuteWithRegions(dSet);

                    doc.Save(templatePath + "Output.pdf");
                    string filename = "Output.pdf"; 
                    Response.Redirect("~/Files/" + filename + "");
                }
            }

        }

        
    }
}