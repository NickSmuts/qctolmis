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
        if (!IsPostBack)
        {
            if (Session["learnerSNum"] != null && Session["learnerSkills"] != null)
            {
                GetLearner(Session["learnerSNum"].ToString());
                GetLearnerQualifications(Session["learnerSNum"].ToString());
                AddSkillsToCV();
                GetClientDetails();
            }
            else
                Response.Redirect("LearnerCVSearch.aspx");

        }
    }

    private void AddSkillsToCV()
    {
        ListItemCollection skillList =(ListItemCollection)HttpContext.Current.Session["learnerSkills"];
        
        if (skillList.Count > 0)
        { 
            lblSkills.Text = "";
            for (int i = 0; i <= skillList.Count - 2; i++)
            {
                lblSkills.Text += skillList[i].ToString() + " , ";
                if (i % 5 ==0 && i>1)
                    lblSkills.Text += "<br/>";
            }
               

            lblSkills.Text += skillList[skillList.Count - 1].ToString();
        }

    }
    private void GetLearnerQualifications(string studentNumber)
    {
        
        try
        {
            lblQualifications.Text = "";
            myConnection.Open();
            string strTemp = @"select DATA.P_Title as title, DATA.FName as fname, DATA.Sname as sname
                ,DATA.ID_NUM as idnum,DATA.Sex as gender,DATA.Student_NUM as snum, DATA.Contact_Cell as cell,
                DATA.Education as education,DATA.Year as eyear,DATA.Marital_Status as mstatus,
                DATA.Disability as disability,
                LearnerData.SCompetent, Standards.SNumber,Standards.Stitle, DATA.Project FROM (DATA INNER JOIN LearnerData ON DATA.Student_NUM = LearnerData.Student_NUM) INNER JOIN Standards ON LearnerData.STitle = Standards.STitle 
                WHERE DATA.Student_NUM = @Student_NUM AND SCompetent ='Competent'";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Parameters.AddWithValue("@Student_NUM", studentNumber);
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();

            while (rdr.Read())
            {
                lblQualifications.Text += rdr["SNumber"].ToString() + " / "+ rdr["Stitle"].ToString() + "<br/>";
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
        
    }
    private void GetClientDetails()
    {
        try
        {

            myConnection.Open();
            string strTemp = @"select CompanyName from Client";
            OleDbCommand myCommand = new OleDbCommand();
           
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();

            while (rdr.Read())
            {
                lblInstitute.Text = rdr["CompanyName"].ToString();
            }
            lblDate.Text = DateTime.Now.ToShortDateString();
            myCommand.Connection.Close();
        }
        catch (Exception er)
        {

        }
        finally
        {
            if (myConnection.State != System.Data.ConnectionState.Closed) myConnection.Close();
        }
    }
    private bool GetLearner(string studentNumber)
    {
        bool result = false;
        try
        {
            
            myConnection.Open();
            string strTemp = @"select DATA.P_Title as title, DATA.FName as fname, DATA.Sname as sname
                ,DATA.ID_NUM as idnum,DATA.Sex as gender,DATA.Student_NUM as snum, DATA.Contact_Cell as cell,
                DATA.Education as education,DATA.Year as eyear,DATA.Marital_Status as mstatus,
                DATA.Disability as disability from DATA
                WHERE DATA.Student_NUM = @Student_NUM";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Parameters.AddWithValue("@Student_NUM", studentNumber);
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();

            while (rdr.Read())
            {
                lblSurname.Text = rdr["sname"].ToString();
                lblName.Text = rdr["fname"].ToString();
                lblGender.Text = rdr["gender"].ToString();
                lblId.Text = rdr["idnum"].ToString();
                lblDisability.Text = rdr["disability"].ToString();
                lblContactNum.Text = rdr["cell"].ToString();
                lblHighestEducation.Text = rdr["education"].ToString();
                lblYear.Text = rdr["eyear"].ToString();
                lblmaritalStatus.Text = rdr["mstatus"].ToString();
                
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


}