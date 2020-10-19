using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace PosgradoWeb
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                PopulateData();
                lblMessage.Text = "Current Database Data!";
            }

        }

        private void PopulateData()
        {
            using (bdposbotEntities dc = new bdposbotEntities())
            {
                gvData.DataSource = dc.Pays.ToList();
                gvData.DataBind();
            }
        }

        protected void btnImport_Click(object sender, EventArgs e)
        {
            if (FileUpload1.PostedFile.ContentType == "application/vnd.ms-excel" ||
                FileUpload1.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                try
                {
                    string fileName = Path.Combine(Server.MapPath("~/ImportDocument"), Guid.NewGuid().ToString() + Path.GetExtension(FileUpload1.PostedFile.FileName));
                    FileUpload1.PostedFile.SaveAs(fileName);

                    string conString = "";
                    string ext = Path.GetExtension(FileUpload1.PostedFile.FileName);

                    if (ext.ToLower() == ".xls")
                    {
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\""; 
                    }
                    else if (ext.ToLower() == ".xlsx")
                    {
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0 Xml;HDR=Yes;IMEX=2\"";
                        //conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }

                    string query = "select [id],[ci],[apellidos],[nombres],[cuotaUno],[cuotaDos],[cuotaTres],[cuotaCuatro],[cuotaCinco],[cuotaSeis],[idCurso] from Pays";
                    OleDbConnection con = new OleDbConnection(conString);
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }

                    OleDbCommand cmd = new OleDbCommand(query, con);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    da.Dispose();
                    con.Close();
                    con.Dispose();

                    //Import to Database
                    using (bdposbotEntities dc = new bdposbotEntities())
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            string id = dr["id"].ToString();
                            var v = dc.Pays.Where(a => a.id.Equals(id)).FirstOrDefault();
                            if (v != null)
                            {
                                //Update here
                                v.ci = dr["ci"].ToString();
                                v.apellidos = dr["apellidos"].ToString();
                                v.nombres = dr["nombres"].ToString();
                                v.cuotaUno = Convert.ToDouble(dr["cuotaUno"]);
                                v.cuotaDos = Convert.ToDouble(dr["cuotaUno"]);
                                v.cuotaTres = Convert.ToDouble(dr["cuotaUno"]);
                                v.cuotaCuatro = Convert.ToDouble(dr["cuotaUno"]);
                                v.cuotaCinco = Convert.ToDouble(dr["cuotaUno"]);
                                v.cuotaSeis = Convert.ToDouble(dr["cuotaUno"]);
                                v.idCurso = dr["idCurso"].ToString();
                            }
                            else
                            {
                                //Insert
                                dc.Pays.Add(new Pays
                                {
                                    id = Convert.ToInt32(dr["cuotaUno"]),
                                    ci = dr["ci"].ToString(),
                                    apellidos = dr["apellidos"].ToString(),
                                    nombres = dr["nombres"].ToString(),
                                    cuotaUno = Convert.ToDouble(dr["cuotaUno"]),
                                    cuotaDos = Convert.ToDouble(dr["cuotaUno"]),
                                    cuotaTres = Convert.ToDouble(dr["cuotaUno"]),
                                    cuotaCuatro = Convert.ToDouble(dr["cuotaUno"]),
                                    cuotaCinco = Convert.ToDouble(dr["cuotaUno"]),
                                    cuotaSeis = Convert.ToDouble(dr["cuotaUno"]),
                                    idCurso = dr["idCurso"].ToString()
                                });
                            }
                        }
                        dc.SaveChanges();
                    }
                    PopulateData();
                    lblMessage.Text = "Successfully data import done!";
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
    }
}