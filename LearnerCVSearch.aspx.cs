using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class LearnerCV : System.Web.UI.Page
{
    OleDbConnection myConnection = new OleDbConnection(ConfigurationManager.AppSettings["connectionstring"].ToString());
    protected void Page_Load(object sender, EventArgs e)
    {
         
    }

    protected void BtnSubmit_Click(object sender, EventArgs e)
    {
        lblError.Text = "";
        if (!string.IsNullOrEmpty(txtStudentNumber.Text))
        {

            if (!GetLearner(txtStudentNumber.Text))
                lblError.Text = "Student Not found!";
            else
            {
                Response.Redirect("LearnerCVSkills.aspx");
            }
                


        }
    }
    
    
    private bool GetLearner(string studentNumber)
    {
        bool result = false;
        try
        {
            
            myConnection.Open();
            string strTemp = "select DATA.P_Title as title, DATA.FName as fname, DATA.Sname as sname, DATA.Student_NUM as snum from DATA WHERE DATA.Student_NUM = @Student_NUM";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Parameters.AddWithValue("@Student_NUM", studentNumber);
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();

            while (rdr.Read())
            {
                Session["learnerSNum"] = rdr["snum"].ToString();
                result = true;
            }
            myCommand.Connection.Close();
        }
        catch (Exception er)
        {

        }
        finally
        {
            if (myConnection.State != System.Data.ConnectionState.Closed) myConnection.Close();
        }
        return result;
    }
    protected void BtnReset_Click(object sender, EventArgs e)
    {

    }

   
    
}