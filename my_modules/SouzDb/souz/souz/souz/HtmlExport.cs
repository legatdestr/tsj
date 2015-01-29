using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace souz.DataExport
    {
    static class HtmlExport
        {
        public static void Export()
            {
            string html = 
                @"<!DOCTYPE html>"
                +"<html>"
                +"<head>"
                +"<meta charset=\"UTF-8\">"
                +"<title>Отчёт</title>"
                +"</head>"
                +"<body>"
                ;
            string footer =
                "</body>"
                + "</html>";
            }
        }
    }
