﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.IO;

public partial class sitemap_chestees : Page
{
  public myFunctions myFunctionsInstance = new myFunctions();
  public constants varConst = new constants();
  protected void Page_Load(object sender, EventArgs e)
  {
    StringBuilder strXML = new StringBuilder();
    strXML.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    strXML.Append("<urlset xmlns=\"http://www.sitemaps.org/schemas/sitemap/0.9\">");
    strXML.Append("<url>");
    strXML.Append("<loc>http://www.chestees.com/</loc>");
    strXML.Append("<changefreq>weekly</changefreq>");
    strXML.Append("</url>");
    strXML.Append("<url>");
    strXML.Append("<loc>http://www.chestees.com/busted-tees/</loc>");
    strXML.Append("<changefreq>weekly</changefreq>");
    strXML.Append("</url>");
    strXML.Append("<url>");
    strXML.Append("<loc>http://www.chestees.com/snorg-tees/</loc>");
    strXML.Append("<changefreq>weekly</changefreq>");
    strXML.Append("</url>");
    strXML.Append("<url>");
    strXML.Append("<loc>http://www.chestees.com/faq-chestees/</loc>");
    strXML.Append("<changefreq>weekly</changefreq>");
    strXML.Append("</url>");
    strXML.Append("<url>");
    strXML.Append("<loc>http://www.chestees.com/contact-chestees/</loc>");
    strXML.Append("<changefreq>weekly</changefreq>");
    strXML.Append("</url>");

    using (varConst.conn)
    {
      varConst.conn.Open();
      SqlCommand cmd = new SqlCommand();
      cmd.Connection = varConst.conn;
      cmd.CommandType = CommandType.StoredProcedure;
      cmd.CommandText = "usp_Digg_Sitemap_Links_Chestees";
      SqlDataReader rdr = cmd.ExecuteReader();
      if (rdr.HasRows)
      {

        while (rdr.Read())
        {
          string Product = Server.UrlEncode(myFunctionsInstance.Stripper(rdr["Title"].ToString()));
          string Slug = rdr["Slug"].ToString();
          int DiggID = Convert.ToInt32(rdr["DiggID"]);
          int DiggStoreID = Convert.ToInt32(rdr["DiggStoreID"]);
          if (rdr["ProductID"] != DBNull.Value) {
            int ProductID = Convert.ToInt32(rdr["ProductID"]);
          }

          if(DiggStoreID == 1 || DiggStoreID == 2) {
              strXML.Append("<url>");
              strXML.Append("<loc>http://www.chestees.com/t-shirts/detail/" + DiggID + "/" + Slug + "/</loc>");
              strXML.Append("<changefreq>weekly</changefreq>");
              strXML.Append("</url>");
            } else if(DiggStoreID == 4) {
              if(ProductID > 0) {
                strXML.Append("<url>");
                strXML.Append("<loc>http://www.chestees.com/funny-t-shirts/" + ProductID + "/" + Slug + "/</loc>");
                strXML.Append("<changefreq>weekly</changefreq>");
                strXML.Append("</url>");
                }
            }
        }
      }
      rdr.Close();

      strXML.Append("</urlset>");

      StreamWriter swFromFile = new StreamWriter(@"\\fs1-n02\stor2wc1dfw1\407499\407510\www.damptshirts.com\web\content\chestees.xml");
      swFromFile.Write(strXML);
      swFromFile.Flush();
      swFromFile.Close();		
    }
  }
}