using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using TestUploader.testweb;

namespace TestUploader
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            if (FileUpload1.FileName == "")
            {
                return;
            }
            else if (System.IO.Path.GetExtension(FileUpload1.FileName).ToString() != ".xls" && System.IO.Path.GetExtension(FileUpload1.FileName).ToString() != ".xlsx")
            {
                Response.Write("<script type='text/javascript'>alert('Please Upload A Valid Excel File');</script>");
                return;
            }
            else
            {
                Label1.Text = "Please Wait...Copying Data";
                Service1 ser = new Service1();
                string path = SaveFile(FileUpload1.PostedFile);//System.IO.Path.GetFullPath(FileUpload1.FileName);
                string s = ser.upload(path);
                Label1.Text = s;
            }
        }
        
        string SaveFile(HttpPostedFile file)
        {
            // Specify the path to save the uploaded file to.
            string savePath = "c:\\Apps\\Uploads\\";

            // Get the name of the file to upload.
            string fileName = FileUpload1.FileName;

            // Create the path and file name to check for duplicates.
            string pathToCheck = savePath + fileName;

            // Create a temporary file name to use for checking duplicates.
            string tempfileName = "";

            // Check to see if a file already exists with the
            // same name as the file to upload.        
            if (System.IO.File.Exists(pathToCheck))
            {
                int counter = 2;
                while (System.IO.File.Exists(pathToCheck))
                {
                    // if a file with this name already exists,
                    // prefix the filename with a number.
                    tempfileName = System.IO.Path.GetFileNameWithoutExtension(fileName) + counter.ToString() + System.IO.Path.GetExtension(fileName);
                    pathToCheck = savePath + tempfileName;
                    counter++;
                }

                fileName = tempfileName;

                // Notify the user that the file name was changed.
            }
            else
            {
                // Notify the user that the file was saved successfully.
                
            }

            // Append the name of the file to upload to the path.
            savePath += fileName;

            // Call the SaveAs method to save the uploaded
            // file to the specified directory.
            FileUpload1.SaveAs(savePath);
            return savePath;

        }

    }
}