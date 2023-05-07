using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SQL_MANIPULATION
{
    public partial class Form1 : Form
    {
        private OleDbConnection thisConnection;


        public Form1()
        {
            InitializeComponent();
            thisConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\Klein\\Documents\\KOMPANYA.accdb");

        }

        private void showEmployee_Click(object sender, EventArgs e)
        {


            string sql = "SELECT EMPLOYEEIDNO, EMPLOYEEFIRSTNAME + ' ' + EMPLOYEELASTNAME as FULLNAME, EMPLOYEEDEPARTMENT, EMPLOYEESALARY FROM EMPLOYEEFILE WHERE EMPLOYEESALARY > (SELECT AVG(EMPLOYEESALARY) FROM EMPLOYEEFILE WHERE EMPLOYEEDEPARTMENT = 'PRODUCTION')";
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, thisConnection);
            DataTable dataTable = new DataTable();
            thisConnection.Open();
            adapter.Fill(dataTable);
            grid1.DataSource = dataTable;
            thisConnection.Close();


        }

    

        private void ttlSalary_Click(object sender, EventArgs e)
        {
    
            string sql = "SELECT SUM(EMPLOYEESALARY) FROM EMPLOYEEFILE";

                OleDbCommand command = new OleDbCommand(sql, thisConnection);
                thisConnection.Open();
                object result = command.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    decimal totalSalaries = Convert.ToInt32(result);
                    MessageBox.Show(string.Format("Total Salary = {0:C}", totalSalaries));
                }
                thisConnection.Close();
             
            }

        private void AvgSalaryLnameM_Click(object sender, EventArgs e)
        {
            string sql = "SELECT AVG(EMPLOYEESALARY) FROM EMPLOYEEFILE WHERE EMPLOYEELASTNAME LIKE 'M%'";

            OleDbCommand command = new OleDbCommand(sql, thisConnection);
            thisConnection.Open();
            object result = command.ExecuteScalar();
            if (result != null && result != DBNull.Value)
            {
                decimal avgSalaryStartsM = Convert.ToInt32(result);
                MessageBox.Show(string.Format("Average Salary of Employees With Lastname That Starts With Letter M = {0:C}", avgSalaryStartsM));
            }
            thisConnection.Close();

        }

        private void DisplaySalaryGreaterthanAVGSalaryWithFullName_Click(object sender, EventArgs e)
        {

            string sql = "SELECT EMPLOYEEIDNO as [ID NUMBER], EMPLOYEEFIRSTNAME as FIRSTNAME, EMPLOYEELASTNAME as LASTNAME, EMPLOYEESALARY as SALARY FROM EMPLOYEEFILE WHERE EMPLOYEEDEPARTMENT = 'PRODUCTION' AND EMPLOYEESALARY > (SELECT AVG(EMPLOYEESALARY) FROM EMPLOYEEFILE WHERE EMPLOYEEDEPARTMENT = 'PRODUCTION')";
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, thisConnection);
            DataTable dataTable = new DataTable();
            thisConnection.Open();
            adapter.Fill(dataTable);
            grid1.DataSource = dataTable;
            thisConnection.Close();
        }

        private void DisplayttlSalaryofAllSales_Click(object sender, EventArgs e)
        {
            string sql = "SELECT SUM(EMPLOYEESALARY) FROM EMPLOYEEFILE WHERE EMPLOYEEDEPARTMENT = 'SALES'";

            OleDbCommand command = new OleDbCommand(sql, thisConnection);
            thisConnection.Open();
            object result = command.ExecuteScalar();
            if (result != null && result != DBNull.Value)
            {
                int totalSalesDepartment = Convert.ToInt32(result);
                MessageBox.Show(string.Format("Total Sales Department = {0:C}", totalSalesDepartment));
            }
            thisConnection.Close();
        }

        private void DisplayPaidLeast_Click(object sender, EventArgs e)
        {
            string sql = "SELECT EMPLOYEEIDNO, EMPLOYEELASTNAME, EMPLOYEEDEPARTMENT FROM EMPLOYEEFILE";
         
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, thisConnection);
            DataTable dataTable = new DataTable();
            thisConnection.Open();
            adapter.Fill(dataTable);
            grid1.DataSource = dataTable;
            thisConnection.Close();


        }

        private void DisplayAllEmployeesNameDepartmentTrainingCourseAttended_Click(object sender, EventArgs e)
        {
            string sql = "SELECT EMPLOYEEFIRSTNAME, EMPLOYEELASTNAME, EMPLOYEEDEPARTMENT, TRAININGCOURSE\r\nFROM ATTENDANCEFILE\r\nINNER JOIN EMPLOYEEFILE ON ATTENDANCEFILE.ATTENDTRAININGEMPLOYEEID = EMPLOYEEFILE.EMPLOYEEIDNO\r\nINNER JOIN TRAININGFILE ON ATTENDANCEFILE.ATTENDTRAININGCODE = TRAININGFILE.TRAININGCODE";




            OleDbDataAdapter adapter = new OleDbDataAdapter(sql,thisConnection);
      
            DataTable dataTable = new DataTable();
            thisConnection.Open();
            adapter.Fill(dataTable);
            DisplayAllNameDepartmentTrainingCourseAttended();
            grid1.DataSource = dataTable;
        
            thisConnection.Close();

        }
        private void DisplayAllNameDepartmentTrainingCourseAttended()
        {
            string sql = "SELECT TRAININGCOURSE FROM TRAININGFILE";



            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, thisConnection);

            DataTable dataTable = new DataTable();
        
            adapter.Fill(dataTable);
            grid1.DataSource = dataTable;
   
        } 

        private void EMPLOYEEMORETHANONETRAINING_Click(object sender, EventArgs e)
        {

            string sql = "SELECT EMPLOYEELASTNAME AS [LASTNAME], COUNT(*) AS [COUNTER] FROM ATTENDANCEFILE INNER JOIN EMPLOYEEFILE ON ATTENDANCEFILE.ATTENDTRAININGEMPLOYEEID = EMPLOYEEFILE.EMPLOYEEIDNO GROUP BY EMPLOYEELASTNAME HAVING COUNT(*) >= 1";
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, thisConnection);
            DataTable dataTable = new DataTable();
            thisConnection.Open();
            adapter.Fill(dataTable);
            grid1.DataSource = dataTable;
            thisConnection.Close();
        }

        private void DisplayAngerMangementTrainingCourse_Click(object sender, EventArgs e)
        {
            string sql = "SELECT EMPLOYEELASTNAME, EMPLOYEEFIRSTNAME, EMPLOYEEDEPARTMENT FROM EMPLOYEEFILE WHERE EMPLOYEEIDNO IN (SELECT ATTENDTRAININGEMPLOYEEID FROM ATTENDANCEFILE  WHERE ATTENDTRAININGCODE = 'TRAINING202')";
            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, thisConnection);
            DataTable dataTable = new DataTable();
            thisConnection.Open();
            adapter.Fill(dataTable);
            grid1.DataSource = dataTable;
            thisConnection.Close();

        }
    }
    
}

