using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class AddSkill : System.Web.UI.Page
{
    OleDbConnection myConnection = new OleDbConnection(ConfigurationManager.AppSettings["connectionstring"].ToString());
    protected void Page_Load(object sender, EventArgs e)
    {
        AddTableToDBIfNotExist();
        DisplaySkills();
    }

    protected void BtnUpdate_Click(object sender, EventArgs e)
    {
        lblError.Text = "";
        if (!string.IsNullOrEmpty(txtSkill.Text))
        {
            if (!CheckIfSkillExists(txtSkill.Text))
                AddSkills(txtSkill.Text);
            else
                lblError.Text = "Skill Already Exists!";
        }
    }

    private void AddTableToDBIfNotExist()
    {
        try
        {
            myConnection.Open();
            string strTemp = " [Skill] Text";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = "CREATE TABLE Skill(" + strTemp + ")";
            myCommand.ExecuteNonQuery();
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

    private bool CheckIfSkillExists(string skill)
    {
        bool result = false;
        try
        {
           
            myConnection.Open();
            string strTemp = " select * from Skill where skill=@skill";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Parameters.AddWithValue("@skillname", skill);
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();
            
            while (rdr.Read())
            {
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
    private void DisplaySkills()
    {
        try
        {
            lblSkills.Text = "";
            myConnection.Open();
            string strTemp = " select * from Skill";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();
           
            while (rdr.Read())
            {
                lblSkills.Text += "<tr><td>" + rdr[0].ToString() + "</td></tr>";
                
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
    private void AddSkills(string skill)
    {
        
        try
        {
            lblError.Text = "";
            myConnection.Open();
            string strTemp = " insert into Skill values (@skillname)";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            myCommand.Parameters.AddWithValue("@skillname", skill);
            myCommand.ExecuteNonQuery();
            myCommand.Connection.Close();
            txtSkill.Text = "";
            DisplaySkills();
        }
        catch (Exception er)
        {
            lblError.Text = er.Message;
        }
        finally
        {
            if (myConnection.State != System.Data.ConnectionState.Closed) myConnection.Close();
        }
    }
}