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

                object replaceAll = WdReplace.wdReplaceAll;
                findAndReplace.replace(wordApp, "<Title>", txtTitle.Text);
                findAndReplace.replace(wordApp, "<name>", txtName.Text);

                //edit Header and footer
                foreach (Section section in doc.Sections)
                {
                    //find the footer and replace
                    Range footerRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Find.Text = "<Date>";
                    footerRange.Find.Replacement.Text = DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss").ToString();
                    footerRange.Find.Execute(ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                    //find the Header and replace
                    Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Find.Text = "<CompanyName>";
                    headerRange.Find.Replacement.Text = "冰冰无限公司";
                    headerRange.Find.Execute(ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                }

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

            }


        }
    }
}