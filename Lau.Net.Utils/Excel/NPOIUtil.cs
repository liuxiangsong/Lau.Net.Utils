using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Lau.Net.Utils.Excel.NpoiStaticUtil;
using Lau.Net.Utils.Excel.NpoiExtensions;

namespace Lau.Net.Utils.Excel
{
    public class NpoiUtil
    {
        private IWorkbook _workbook;
        public IWorkbook Workbook
        {
            get { return this._workbook; }
            set { this._workbook = value; }
        }
        public NpoiUtil(ExcelType type = ExcelType.Xls)
        {
            this.Workbook = NpoiStaticUtil.CreateWorkbook(type);
        }

        /// <summary>
        /// 将DataTable添加至Workbook中
        /// </summary>
        /// <param name="sourceTable">源数据表</param>
        /// <param name="workbook">目标Workbook</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="startRow">导出到Excel中的起始行</param>
        public ISheet DataTableToWorkbook(DataTable sourceTable, bool isExportCaption = true, int startRow = 0)
        {
            return NpoiStaticUtil.DataTableToWorkbook(sourceTable, this.Workbook, isExportCaption, startRow);
        }


        ///// <summary>
        ///// 通过HttpContext导出excel
        ///// </summary>
        ///// <param name="fileName">excel文件名</param>
        //public void ExportByHttpContext(string fileName)
        //{
        //    System.Web.HttpContext curContext = System.Web.HttpContext.Current;
        //    // 设置编码和附件格式
        //    curContext.Response.ContentType = "application/ms-excel";
        //    curContext.Response.ContentEncoding = Encoding.UTF8;
        //    curContext.Response.Charset = "";
        //    //------------------
        //    //这里判断使用的浏览器是否为Firefox，Firefox导出文件时不需要对文件名显示编码，编码后文件名会乱码
        //    //但是IE和Google需要编码才能保持文件名正常
        //    if (curContext.Request.ServerVariables["http_user_agent"].ToString().IndexOf("Firefox") != -1)
        //    {
        //        curContext.Response.AddHeader("Content-Disposition", "attachment;filename=" + fileName);
        //    }
        //    else
        //    {
        //        curContext.Response.AddHeader("Content-Disposition", "attachment;filename="
        //                                                             + System.Web.HttpUtility.UrlEncode(fileName,
        //                                                                 System.Text.Encoding.UTF8));
        //    }

        //    using (MemoryStream ms = new MemoryStream())
        //    {
        //        this.Workbook.Write(ms);
        //        ms.Flush();
        //        ms.Position = 0;
        //        curContext.Response.BinaryWrite(ms.GetBuffer());
        //        curContext.Response.End();
        //    }
        //}

    }



}
