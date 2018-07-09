using System;

using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using ToDoc.model;

namespace ToDoc
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Unnamed_Click(object sender, EventArgs e)
        {
            createWordDocument(Server.MapPath("Example.docx"));
        }
        private void createWordDocument(object filename)
        {
            object missing = Missing.Value;
            Application wordApp = new Application();
            Document doc = null;

            //check the template is exist or not
            if (File.Exists((string)filename))
            {
                FindAndReplace findAndReplace = new FindAndReplace();
                object readOnly = false; //default
                object isVisible = false;//make the doc visible

                wordApp.Visible = false;

                //open the template
                doc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);


                //create a new doc same as the template
                doc.Activate();
                //insert the value to the new doc follow the id
                findAndReplace.replace(wordApp, "<Title>", txtTitle.Text);
                findAndReplace.replace(wordApp, "<name>", txtName.Text);

                object tempFile = Server.MapPath("Temp/" + DateTime.Now.ToString("hhmmssffffff") + ".docx");
                //save the new doc in temp file
                doc.SaveAs2(ref tempFile, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);
                //close all the doc
                doc.Close(ref missing, ref missing, ref missing);

                //let the user download the doc
                Response.ContentType = "Application/msword";
                Response.AddHeader("Content-Disposition", "attachment;filename=" + tempFile);
                Response.TransmitFile(Path.Combine(tempFile.ToString()));
                Response.End();

                //delete the doc genenrate ... ***optinal
                File.Delete(tempFile.ToString());

            }


        }
    }
}