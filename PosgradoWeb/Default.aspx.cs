using PosgradoWeb.Controller;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PosgradoWeb
{
    public partial class _Default : Page
    {
        SqlConnection con = new SqlConnection();


        //DataTable dt = new DataTable();

        protected void Page_Load(object sender, EventArgs e)
        {
            con.ConnectionString = "Data Source=bdposbot.database.windows.net;Initial Catalog=bdposbot;User ID=posgradobot;Password=Posbot123;MultipleActiveResultSets=True;Application Name=EntityFramework";
            con.Open();
            if (!IsPostBack)
            {
                PopulateData();
                lblMessage.Text = "Current Database Data!";
            }

        }

        private void PopulateData()
        {
            //using (bdEntities dc = new bdEntities())
            //{
            //    gvData.DataSource = dc.Pays.ToList();
            //    gvData.DataBind();
            //}
            //ds = new DataSet();
            //cmd.CommandText = "Select * from Pays";
            //cmd.Connection = con;
            //sda = new SqlDataAdapter(cmd);
            //sda.Fill(ds);
            //cmd.ExecuteNonQuery();
            //gvData.DataSource = ds;
            //gvData.DataBind();
            //con.Close();

            using (bdposbotEntities dc = new bdposbotEntities())
            {
                gvData.DataSource = dc.Pays.ToList();
                gvData.DataBind();
            }
        }

        protected void btnImport_Click(object sender, EventArgs e)
        {
            byte[] xd;
            xd = FileUpload1.FileBytes;

            var namefile = FileUpload1.PostedFile.FileName;
            clsExcel lmao = new clsExcel();
            var list = lmao.mtdConvertirExcel(xd, null);
            //var name = list[1][1].ToString();

            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Select * from Pays";
            cmd.Connection = con;

            SqlDataAdapter sda = new SqlDataAdapter(cmd);

            DataSet ds = new DataSet();
            sda.Fill(ds);
            sda.Dispose();
            con.Close();
            con.Dispose();

            // Import to Database
            using (bdposbotEntities dc = new bdposbotEntities())
            {
                foreach (var ls in list)
                {
                    if(ls[0] != "Id")
                    {
                        string ci = ls[1];
                        var v = dc.Pays.Where(a => a.ci.Equals(ci)).FirstOrDefault();
                        if (v != null)
                        {
                            //Update here
                            v.ci = ls[1];
                            v.apellidos = ls[2];
                            v.nombres = ls[3];
                            v.cuotaUno = Convert.ToDouble(ls[4]);
                            v.cuotaDos = Convert.ToDouble(ls[5]);
                            v.cuotaTres = Convert.ToDouble(ls[6]);
                            v.cuotaCuatro = Convert.ToDouble(ls[7]);
                            v.cuotaCinco = Convert.ToDouble(ls[8]);
                            v.cuotaSeis = Convert.ToDouble(ls[9]);
                            v.idCurso = ls[10];
                        }
                        else
                        {
                            //Insert
                            dc.Pays.Add(new Pays
                            {
                                id = Convert.ToInt32(ls[0]),
                                ci = ls[1],
                                apellidos = ls[2],
                                nombres = ls[3],
                                cuotaUno = Convert.ToDouble(ls[4]),
                                cuotaDos = Convert.ToDouble(ls[5]),
                                cuotaTres = Convert.ToDouble(ls[6]),
                                cuotaCuatro = Convert.ToDouble(ls[7]),
                                cuotaCinco = Convert.ToDouble(ls[8]),
                                cuotaSeis = Convert.ToDouble(ls[9]),
                                idCurso = ls[10]
                            });
                        }
                    }
                }
                dc.SaveChanges();
            }
            PopulateData();
            lblMessage.Text = "Successfully data import done!";
        }
    }
}