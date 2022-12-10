namespace Invoicing.model
{
    /// <summary>
    /// 整張儲存
    /// </summary>
    internal class save_list
    {
        /// <summary>
        /// 單子編號
        /// </summary>
        public string bili { get; set; }
        /// <summary>
        /// 貨品編號
        /// </summary>
        public string store { get; set; }
        /// <summary>
        /// 品名
        /// </summary>
        public string name_total { get; set; }
        /// <summary>
        /// 數量
        /// </summary>
        public string number { get; set; }
        /// <summary>
        /// 基本單位
        /// </summary>
        public string unit { get; set; }
        /// <summary>
        /// 單價
        /// </summary>
        public string unit_cost { get; set; }
        /// <summary>
        /// 金額
        /// </summary>
        public string cost { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string remark { get; set; }
    }
    /// <summary>
    /// 採購單
    /// </summary>
    internal class purchase_save
    {
        /// <summary>
        /// 單子編號
        /// </summary>
        public string bili { get; set; }
        /// <summary>
        /// 貨品編號
        /// </summary>
        public string store { get; set; }
        /// <summary>
        /// 品名
        /// </summary>
        public string name_total { get; set; }
        /// <summary>
        /// 數量
        /// </summary>
        public string number { get; set; }
        /// <summary>
        /// 基本單位
        /// </summary>
        public string unit { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string remark { get; set; }
    }
}
