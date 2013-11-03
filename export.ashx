<%@ WebHandler Language="C#" Class="export" %>

using System;
using System.Web;

public class export : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        //context.Response.ContentType = "text/plain";
        context.Response.Write("Hello World");

        context.Response.AddHeader("Content-Type", "application/force-download");
        context.Response.AddHeader("Content-Type", "application/vnd.ms-excel");
        context.Response.AddHeader("Content-Disposition", "attachment;filename=export.xls");
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}