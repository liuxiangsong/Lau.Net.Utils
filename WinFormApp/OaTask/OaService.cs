using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lau.Net.Utils;
using Microsoft.Extensions.Logging;

namespace WinFormApp.OaTask
{
    public class OaService
    {
        private readonly IFreeSql _freeSql;
        private readonly ILogger<OaService> _logger;
        public OaService(IFreeSql freeSql, ILogger<OaService> logger)
        {
            this._freeSql = freeSql;
            _logger = logger;
        }

        //public DataTable Query()
        //{
        //  return  _freeSql.Ado.ExecuteDataTable(" SELECT * FROM LR_Base_User WHERE F_RealName='刘鸿森'");
        //}

        /// <summary>
        /// 查询是否有异常的PS任务
        /// </summary>
        /// <returns></returns>
        public async Task<bool> ExistsAbnormalTask()
        {
            var dt = await GetAbnormalTaskAsync();
            if (dt == null || dt.Rows.Count < 1)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 校正当天异常任务
        /// </summary>
        public async void RectifyAbnormalTask()
        {
            var dt = await GetAbnormalTaskAsync();
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            var sb = new StringBuilder();
            foreach(DataRow row in dt.Rows)
            {
                sb.Append($"UPDATE MP_XZ_ryrw SET ");
                var IsBeginAbnormal = row["IsBeginAbnormal"].ToString() == "1";
                if (IsBeginAbnormal)
                {
                    var beginTime = row.GetValue<DateTime>("计划开始时间").AddMinutes(2).ToString("yyyy-MM-dd HH:mm:ss");
                    sb.Append($" sjkssj = '{beginTime}'");
                }
                if (row["IsEndAbnormal"].ToString() == "1")  
                {
                    if (IsBeginAbnormal)
                    {
                        sb.Append(",");
                    }
                    var endTime = row.GetValue<DateTime>("计划结束时间").AddMinutes(-2).ToString("yyyy-MM-dd HH:mm:ss");
                    sb.Append($" sjjssj = '{endTime}',yqjsyy=null");
                }
                sb.AppendLine($" WHERE id='{row.GetValue<string>("id")}'");
            }
            _logger.LogInformation(sb.ToString());
            await _freeSql.Ado.ExecuteNonQueryAsync(sb.ToString());
        }

        /// <summary>
        /// 获取当天异常任务（开始时间晚于计划开始时间10分钟，或结束时间早于计划结束时间10分钟以上)
        /// </summary>
        /// <returns></returns>
        public async Task<DataTable> GetAbnormalTaskAsync( )
        {
            var sql = $@"
SELECT bm,id,rwkssj '计划开始时间',rwjssj '计划结束时间', sjkssj '实际开始时间',sjjssj '实际结束时间',yqjsyy,
CASE WHEN sjkssj >  DATEADD(MINUTE, 10, rwkssj) THEN 1 ELSE 0 END 'IsBeginAbnormal',
CASE WHEN DATEADD(MINUTE, 10, sjjssj) < rwjssj OR sjjssj > rwjssj THEN 1 ELSE 0 END 'IsEndAbnormal'
FROM MP_XZ_ryrw WHERE  zpryxm='03401127-767D-4493-A292-F333F38C93F5' 
AND CONVERT(date, rwkssj)  = CONVERT(date, GETDATE())
AND (sjkssj >  DATEADD(MINUTE, 10, rwkssj) OR DATEADD(MINUTE, 10, sjjssj) < rwjssj OR sjjssj > rwjssj)
";
            var dt = await _freeSql.Ado.ExecuteDataTableAsync(sql);
            return dt;
        }

        public DataTable GetTodayTask()
        {
            var sql = $@" 
SELECT bm,id,rwkssj '计划开始时间',rwjssj '计划结束时间', sjkssj '实际开始时间',sjjssj '实际结束时间',yqjsyy,* 
FROM MP_XZ_ryrw WHERE zpryxm='03401127-767D-4493-A292-F333F38C93F5' 
AND CONVERT(date, rwkssj)  = CONVERT(date, GETDATE())
ORDER BY rwkssj DESC
";
            return _freeSql.Ado.ExecuteDataTable(sql);
        }

        public bool FinishTodayTask()
        {
            var sql = @"UPDATE LR_NWF_Process SET F_IsFinished=1,F_FinishedDate=GETDATE() 
WHERE F_IsFinished = 0 and EXISTS(SELECT bm FROM MP_XZ_ryrw WHERE zpryxm='03401127-767D-4493-A292-F333F38C93F5' 
AND CONVERT(date, rwkssj)  = CONVERT(date, GETDATE()) AND F_BillNo= bm)";
            return _freeSql.Ado.ExecuteNonQuery(sql)>0;
        }
    }
}
