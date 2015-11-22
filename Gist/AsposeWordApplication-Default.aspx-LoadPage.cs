// For complete examples and data files, please go to https://github.com/atirtahir/PDF-invoice-generator-using-aspose.word-api

            //new code
            //another

            if (!Page.IsPostBack)
            {
                List<int> lst = new List<int>();
                lst.Add(1);
                lst.Add(2);
                lst.Add(3);
                ItemDetails.DataSource = lst;
                ItemDetails.DataBind();
            }
