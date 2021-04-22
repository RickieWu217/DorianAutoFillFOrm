using System;
using System.Collections.Generic;
using System.Text;

namespace SeleniumExample
{
    public class ExcelDto
    {
        [ColumnIndex(1)]
        public string 合約起始日 { get; set; }
        [ColumnIndex(2)]
        public string 合約迄止日 { get; set; }
        [ColumnIndex(3)]
        public string 學校合約編號 { get; set; }
        [ColumnIndex(4)]        
        public string 是否公開 { get; set; }
        [ColumnIndex(5)] 
        public string 合約名稱 { get; set; }
        [ColumnIndex(6)] 
        public string 是否跨校 { get; set; }
        [ColumnIndex(7)] 
        public string 資金來源 { get; set; }
        [ColumnIndex(8)] 
        public string 合作模式 { get; set; }
        [ColumnIndex(9)] 
        public string 合約金額來源_來自政府 { get; set; }
        [ColumnIndex(10)] 
        public string 合約金額來源_來自企業 { get; set; }
        [ColumnIndex(11)] 
        public string 合約金額來源_來自其他單位 { get; set; }
        [ColumnIndex(12)] 
        public string 合約總金額 { get; set; }
        [ColumnIndex(13)] 
        public string 上傳合約書影本 { get; set; }
        [ColumnIndex(14)] 
        public string 本案分潤至學校聯盟_預定分潤金額 { get; set; }
        [ColumnIndex(15)] 
        public string 若分潤金額為0_請說明原因 { get; set; }
        [ColumnIndex(16)] 
        public string 平台協助項目_主要 { get; set; }
        [ColumnIndex(17)] 
        public string 平台協助項目_次要 { get; set; }
        [ColumnIndex(18)] 
        public string 請說明平台具體加值內容與協助項目 { get; set; }

    }
}
