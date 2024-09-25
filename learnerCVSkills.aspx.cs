using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class learnerCVSkills : System.Web.UI.Page
{
    OleDbConnection myConnection = new OleDbConnection(ConfigurationManager.AppSettings["connectionstring"].ToString());
    List<string> cvSkills = new List<string>();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            AddTableToDBIfNotExist();
            if (Session["learnerSNum"] != null)
            {
                cvSkills.Clear();
                GetAllSkills();
                GetLearner(Session["learnerSNum"].ToString());
                GetLearnerSkills(Session["learnerSNum"].ToString());
                
            }
            else
                Response.Redirect("LearnerCVSearch.aspx");
            
        }
    }
    
    private bool GetLearner(string studentNumber)
    {
        bool result = false;
        try
        {
            learnerPanel.Visible = false;
            myConnection.Open();
            string strTemp = "select DATA.P_Title as title, DATA.FName as fname, DATA.Sname as sname, DATA.Student_NUM as snum from DATA WHERE DATA.Student_NUM = @Student_NUM";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Parameters.AddWithValue("@Student_NUM", studentNumber);
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();

            while (rdr.Read())
            {
                lblSName.Text = rdr["title"].ToString() + " " + rdr["fname"].ToString() + " " + rdr["sname"].ToString();
                lblSNum.Text = rdr["snum"].ToString();
                learnerPanel.Visible = true;
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

    private void GetAllSkills()
    {
        try
        {
            DDSkills.Items.Clear();
            DDSkills.Items.Add(new ListItem("", ""));
            learnerPanel.Visible = false;
            myConnection.Open();
            string strTemp = "select * from Skill";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();

            while (rdr.Read())
            {
                DDSkills.Items.Add(new ListItem(rdr[0].ToString(), rdr[0].ToString(), true));
            }
            myCommand.Connection.Close();
            learnerPanel.Visible = true;
        }
        catch (Exception er)
        {

        }
        finally
        {
            if (myConnection.State != System.Data.ConnectionState.Closed) myConnection.Close();
        }
    }

    private bool GetLearnerSkills(string studentNumber)
    {
        bool result = false;
        try
        {
            
            myConnection.Open();
            string strTemp = "select l.skill from LearnerSkill as l WHERE l.sNum = @Student_NUM";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Parameters.AddWithValue("@Student_NUM", studentNumber);
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            OleDbDataReader rdr = myCommand.ExecuteReader();

            while (rdr.Read())
            {
                skillListBox.Items.Add( new ListItem ( rdr["skill"].ToString(),rdr["skill"].ToString() ));
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

    private void AddTableToDBIfNotExist()
    {
        try
        {
            myConnection.Open();
            string strTemp = " [sNum] Text, [skill] text";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = "CREATE TABLE LearnerSkill(" + strTemp + ")";
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

    private void AddSkillToLearner(string skill,string sNum)
    {

        try
        {  
            myConnection.Open();
            string strTemp = " insert into LearnerSkill (sNum,skill) values (@sNum, @skillname)";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            myCommand.Parameters.AddWithValue("@sNum",sNum);
            myCommand.Parameters.AddWithValue("@skillname", skill);
            myCommand.ExecuteNonQuery();
            myCommand.Connection.Close();
            //txtSkill.Text = "";
            //DisplaySkills();
        }
        catch (Exception er)
        {
           // lblError.Text = er.Message;
        }
        finally
        {
            if (myConnection.State != System.Data.ConnectionState.Closed) myConnection.Close();
        }
    }

    private void RemoveSkillFromLearner(string skill, string sNum)
    {

        try
        {
            myConnection.Open();
            string strTemp = " delete from LearnerSkill where sNum=@sNum and skill=@skillname";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = strTemp;
            myCommand.Parameters.AddWithValue("@sNum", sNum);
            myCommand.Parameters.AddWithValue("@skillname", skill);
            myCommand.ExecuteNonQuery();
            myCommand.Connection.Close();
            //txtSkill.Text = "";
            //DisplaySkills();
        }
        catch (Exception er)
        {
            // lblError.Text = er.Message;
        }
        finally
        {
            if (myConnection.State != System.Data.ConnectionState.Closed) myConnection.Close();
        }
    }

    protected void BtnAddSkill_Click(object sender, EventArgs e)
    {
        if (DDSkills.SelectedItem.Text != "")
            skillListBox.Items.Add(new ListItem(DDSkills.SelectedItem.Text + " : " + txtSkillYears.Text, DDSkills.SelectedItem.Value));
        else
            skillListBox.Items.Add(new ListItem(txtSkillYears.Text, DDSkills.SelectedItem.Value));
        
        if (DDSkills.SelectedItem.Text != "")
            AddSkillToLearner(DDSkills.SelectedItem.Text + " : " + txtSkillYears.Text, Session["learnerSNum"].ToString());
        else
            AddSkillToLearner(txtSkillYears.Text, Session["learnerSNum"].ToString());
        txtSkillYears.Text = "";
        DDSkills.SelectedIndex = 0;
    }

    protected void skillListBox_SelectedIndexChanged(object sender, EventArgs e)
    {
        //try
        //{
        //    DDSkills.SelectedItem.Text = skillListBox.SelectedItem.Text.Split(':')[0].Trim().ToString();
        //    txtSkillYears.Text = skillListBox.SelectedItem.Text.Split(':')[1].ToString();
        //}
        //catch(Exception er)
        //{
        //    DDSkills.SelectedItem.Text = "";
        //    txtSkillYears.Text = skillListBox.SelectedItem.Text;
        //}
       
        

    }

    protected void BtnDeleteSkill_Click(object sender, EventArgs e)
    {
        try
        {
            cvSkills.Remove(skillListBox.SelectedItem.Text);
            RemoveSkillFromLearner(skillListBox.SelectedItem.Text, Session["learnerSNum"].ToString());
            skillListBox.Items.RemoveAt(skillListBox.SelectedIndex);
            txtSkillYears.Text = "";
            DDSkills.SelectedIndex = 0;
        }
        catch (Exception er)
        {

        }

    }

    protected void BtnSubmit_Click(object sender, EventArgs e)
    {
        //Session["learnerSkills"] = cvSkills;
        //foreach (string t in skillListBox.Items)
        //    cvSkills.Add(t.ToString());
        HttpContext.Current.Session.Add("learnerSkills", skillListBox.Items);
        Response.Redirect("LearnerCV.aspx");
    }

    protected void BtnReset_Click(object sender, EventArgs e)
    {
        Response.Redirect("LearnerCVSkills.aspx");
    }
}