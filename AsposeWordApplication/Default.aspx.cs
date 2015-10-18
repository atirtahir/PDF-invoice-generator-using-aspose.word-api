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

            //path where doc will be saved
            var path = Server.MapPath("~/Files/");

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            //company info
            string header = "ASPOSE.WORDS INVOICE";
            string companyOffile = "Cubatore 1e, Park Road";
            string companyAddress = "Islamabad";

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
            builder.Writeln(header);
            builder.Writeln(companyOffile);
            builder.Writeln(companyAddress);


            //pick current date time

            var currentDate = DateTime.Today.ToString("dd-MM-yyyy");

            //client's information
            builder.StartTable();
            builder.InsertCell();
            builder.InsertField(FieldType.FieldPage, false);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.InsertBreak(BreakType.LineBreak);
            builder.Writeln("To: Mr. " + cname.Value);
            builder.InsertBreak(BreakType.LineBreak);
            builder.Writeln("Product Information: " + pinfo.Value);
            builder.InsertBreak(BreakType.LineBreak);
            builder.Writeln("Product Price: " + price.Value);
            builder.InsertBreak(BreakType.LineBreak);
            builder.Writeln("Date: " + currentDate);
            builder.EndTable();

            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Writeln("Dear " + cname.Value);
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Writeln(msg.Value);
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Writeln("Thanks");
            builder.Writeln("Project Manager");

            //save doc
            doc.Save(path + "Invoice.pdf");

        }
    }
}