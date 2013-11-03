<%@ WebHandler Language="C#" Class="test" %>

using System;
using System.Web;
//using Microsoft.Office.Interop;
//using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
public class test : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        context.Response.ContentType = "text/plain";
        context.Response.Write("Hello World");

        Excel.Application excel = new Excel.Application();
        excel.Application.Workbooks.Add(true);
        
        excel.Cells[ 1 , 1 ] = "First Row First Column" ;
        excel.Cells[ 1 , 2 ] = "First Row Second Column" ;
        excel.Cells[ 2 , 1 ] = "Second Row First Column" ;
        excel.Cells[2, 2] = "Second Row Second Column";
        
        excel.Visible = true;
    }
 
    public bool IsReusable {
        get {
            return false;
        }
        
    }

}