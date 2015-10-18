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

namespace AsposeWordApplication
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //path where doc will be saved
            var path = Server.MapPath("~/Files/");

            //pick current date time
            var currentDate = DateTime.Today.ToString("dd-MM-yyyy");
             
            //upload company logo
            if (Request.Files.Count > 0)
            {
                var file2 = Request.Files[0];

                if (file2 != null && file2.ContentLength > 0)
                {

                    var fileName = Path.GetFileName(file2.FileName);
                    var pathOfLogo = Path.Combine(Server.MapPath("~/App_Data/"), fileName);
                    file2.SaveAs(pathOfLogo);
                    builder.InsertBreak(BreakType.LineBreak);

                    //upload logo to the doc
                    System.Drawing.Image image = System.Drawing.Image.FromFile(pathOfLogo);
                    builder.MoveToDocumentStart();
                    builder.InsertBreak(BreakType.LineBreak);
                    builder.PageSetup.Orientation = Aspose.Words.Orientation.Portrait;
                    builder.InsertImage(image);
                }
            }

            //show company's information
            builder.InsertBreak(BreakType.LineBreak);
            builder.InsertBreak(BreakType.LineBreak);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Writeln(comname.Value);
            builder.Writeln(compmail.Value);
            builder.Writeln(add.Value); 

            //client's information
            
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left; 
            builder.Writeln("To: Mr. " + cname.Value);
            builder.Writeln(mail.Value); 
            builder.Writeln("Date: " + currentDate); 

            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Writeln("Dear " + cname.Value+",");
            builder.Writeln("Following are your item(s) details.");
            builder.InsertBreak(BreakType.ParagraphBreak);
            //item(s) description
            builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Item(s)");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln(item1.Value);
            builder.InsertCell();
            builder.Writeln(qty1.Value);
            builder.InsertCell();
            builder.Writeln(price1.Value);
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln(item2.Value);
            builder.InsertCell();
            builder.Writeln(qty2.Value);
            builder.InsertCell();
            builder.Writeln(price2.Value);
            builder.EndRow();
            builder.InsertCell();
            builder.Writeln(item3.Value);
            builder.InsertCell();
            builder.Writeln(qty3.Value);
            builder.InsertCell();
            builder.Writeln(price3.Value);
            builder.EndRow();

            builder.EndTable();
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Writeln("Thanks");
            builder.Writeln("Project Manager");
            doc.Save(path + "Invoice.pdf");
            //test code
            string filename = "Invoice.pdf";
            //Response.AppendHeader("Content-Disposition", "attachment; filename=" + filename + "");
            Response.Redirect("~/Files/"+filename+"");
        }
    }
}