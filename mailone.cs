using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Threading;

namespace Mail
{
    public partial class mailone : Form
    {
        #region 全局变量
        public StringBuilder strErrorPayroll10 = new StringBuilder();
        public StringBuilder strErrorPayroll15 = new StringBuilder();
        public StringBuilder strErrorMaPayroll15 = new StringBuilder();
        public StringBuilder strErrorYearAwardPayroll = new StringBuilder();



        public List<Payroll10> listPayroll10 = null;
        public List<Payroll15> listPayroll15 = null;
        public List<MaPayroll15> listMaPayroll15 = null;
        public List<YearAwardPayroll> listYearAward = null;


        public bool isErrorPayroll10 = false;
        public bool isErrorPayroll15 = false;
        public bool isErrorMaPayroll15 = false;
        public bool isErrorYearAwardPayroll = false;





        DataTable datatable = new DataTable();

        int total10 = 0;
        int current10 = 0;

        int total15 = 0;
        int current15 = 0;

        int totalMa15 = 0;
        int currentMa15 = 0;

        int totalYearAward = 0;
        int currentYearAward = 0;

        #endregion

        #region 委托：设置text值
        delegate void SetTextBox(string str);
        SetTextBox setpayroll10txtfilename, setpayroll15txtfilename, setmapayroll15txtfilename, setyearawardtxtfilename;
        //10号工资
        private void setPayroll10TxtFileName(string filename)
        {
            txtPayroll10.Text = filename;
        }
        //15号员工工资
        private void setPayroll15TxtFileName(string filename)
        {
            txtPayroll15.Text = filename;
        }
        //15号管理者工资
        private void setMaPayroll15TxtFileName(string filename)
        {
            txtMaPayroll15.Text = filename;
        }

        //年终奖工资
        private void setYearAwardTxtFileName(string filename)
        {
            txtYearAward.Text = filename;
        }


        #endregion

        #region 委托：点击事件
        delegate void ButtonClick(object sender, EventArgs e);
        ButtonClick buttonclick1, buttonclick2, buttonclick3, buttonclick4, buttonclick5, buttonclick6, buttonclick7, buttonclick8, buttonclick9;
        #endregion

        #region 委托：按钮是否可使用
        delegate void SetBntEnable(bool b);
        SetBntEnable setbutton1enable, setbutton2enable, setbutton3enable, setbutton4enable, setbutton5enable, setbutton6enable, setbutton7enable, setbutton8enable, setbutton9enable;
        //发送10号
        private void setbutton1Enable(bool b)
        {
            button1.Enabled = b;
        }
        //选取10号工资表
        private void setbutton3Enable(bool b)
        {
            button3.Enabled = b;
        }

        //发送员工15号
        private void setbutton5Enable(bool b)
        {
            button5.Enabled = b;
        }
        //选取员工15号工资表
        private void setbutton4Enable(bool b)
        {
            button4.Enabled = b;
        }

        //发送管理者15号
        private void setbutton7Enable(bool b)
        {
            button7.Enabled = b;
        }
        //选取管理者15号工资表
        private void setbutton6Enable(bool b)
        {
            button6.Enabled = b;
        }

        //发送年终奖
        private void setbutton9Enable(bool b)
        {
            button9.Enabled = b;
        }
        //选取发送年终奖工资表
        private void setbutton8Enable(bool b)
        {
            button8.Enabled = b;
        }

        //退出
        private void setbutton2Enable(bool b)
        {
            button2.Enabled = b;
        }


        #endregion
        public mailone()
        {
            InitializeComponent();

            setpayroll10txtfilename = new SetTextBox(setPayroll10TxtFileName);//10号工资表
            setpayroll15txtfilename = new SetTextBox(setPayroll15TxtFileName);//15号员工工资表
            setmapayroll15txtfilename = new SetTextBox(setMaPayroll15TxtFileName);//15号管理者工资表
            setyearawardtxtfilename = new SetTextBox(setYearAwardTxtFileName);//年终奖表


            button1.Enabled = false;//10号发送按钮
            button5.Enabled = false;//15号发送按钮
            button7.Enabled = false;//15号发送按钮
            button8.Enabled = false;//年终奖发送按钮


            buttonclick3 = new ButtonClick(button3_Click);
            buttonclick1 = new ButtonClick(button1_Click);

            buttonclick4 = new ButtonClick(button4_Click);
            buttonclick5 = new ButtonClick(button5_Click);

            buttonclick6 = new ButtonClick(button6_Click);
            buttonclick7 = new ButtonClick(button7_Click);

            buttonclick8 = new ButtonClick(button8_Click);
            buttonclick9 = new ButtonClick(button9_Click);


            buttonclick2 = new ButtonClick(button2_Click);

            setbutton1enable = new SetBntEnable(setbutton1Enable);
            setbutton2enable = new SetBntEnable(setbutton2Enable);
            setbutton3enable = new SetBntEnable(setbutton3Enable);
            setbutton4enable = new SetBntEnable(setbutton4Enable);
            setbutton5enable = new SetBntEnable(setbutton5Enable);
            setbutton6enable = new SetBntEnable(setbutton6Enable);
            setbutton7enable = new SetBntEnable(setbutton7Enable);
            setbutton8enable = new SetBntEnable(setbutton8Enable);
            setbutton9enable = new SetBntEnable(setbutton9Enable);


        }



        #region 导入10号工资Excel
        private void button3_Click(object sender, EventArgs e)
        {
            Thread threadSelectExcel10 = new Thread(new ThreadStart(SelectExcel10));
            threadSelectExcel10.SetApartmentState(ApartmentState.STA);
            threadSelectExcel10.Start();

        }

        /// <summary>
        /// 读入10号工资Excel
        /// </summary>
        private void SelectExcel10()
        {
            //strErrorScene.Clear();
            OpenFileDialog PayrollFile = new OpenFileDialog();
            PayrollFile.Title = "10号工资Excel文件上传";
            PayrollFile.Filter = "Excel文件|*.xls;*.xlsx";//只选择Excel文件
            if (PayrollFile.ShowDialog() == DialogResult.OK)
            {
                //txtPayroll10.Text = openFileDialog1.FileName;   //得到附件的地址
                listPayroll10 = new List<Payroll10>();
                txtPayroll10.Invoke(setpayroll10txtfilename, new object[] { PayrollFile.FileName });
                DataTable dt = ExcelToDataTable(PayrollFile.FileName, "sheet1");
                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Payroll10 model = new Payroll10();
                        model.ID = dt.Rows[i][0].ToString().Trim();
                        model.Name = dt.Rows[i][1].ToString().Trim();
                        model.Department = dt.Rows[i][2].ToString().Trim();
                        model.Rank = dt.Rows[i][3].ToString().Trim();
                        model.EntryDate = dt.Rows[i][4].ToString().Trim();
                        model.PositiveDate = dt.Rows[i][5].ToString().Trim();
                        model.MonthPlan = dt.Rows[i][6].ToString().Trim();
                        model.PositiveDays = dt.Rows[i][7].ToString().Trim();
                        model.ProbationSalary = dt.Rows[i][8].ToString().Trim();
                        model.PositiveSalary = dt.Rows[i][9].ToString().Trim();
                        model.MonthSalaryStandard = dt.Rows[i][10].ToString().Trim();
                        model.MonthlyFixedSalary = dt.Rows[i][11].ToString().Trim();
                        model.BasePay = dt.Rows[i][12].ToString().Trim();
                        model.BasicSalaryCompensation = dt.Rows[i][13].ToString().Trim();
                        model.MonthlyPerformanceBaseSalary = dt.Rows[i][14].ToString().Trim();
                        model.Seniority = dt.Rows[i][15].ToString().Trim();
                        model.AgeSalary = dt.Rows[i][16].ToString().Trim();
                        model.TravelAllowance = dt.Rows[i][17].ToString().Trim();
                        model.FullAttendenceAward = dt.Rows[i][18].ToString().Trim();
                        model.IncreaseWagesTotal = dt.Rows[i][19].ToString().Trim();
                        model.NotPostDays = dt.Rows[i][20].ToString().Trim();
                        model.NotInCharge = dt.Rows[i][21].ToString().Trim();
                        model.LeaveDays = dt.Rows[i][22].ToString().Trim();
                        model.LeaveFrom = dt.Rows[i][23].ToString().Trim();
                        model.SickLeaveDays = dt.Rows[i][24].ToString().Trim();
                        model.SickLeaveDeductions = dt.Rows[i][25].ToString().Trim();
                        model.LateHourNotPunchAt = dt.Rows[i][26].ToString().Trim();
                        model.LatePayment = dt.Rows[i][27].ToString().Trim();
                        model.OtherDeductions = dt.Rows[i][28].ToString().Trim();
                        model.TotalDeductions = dt.Rows[i][29].ToString().Trim();
                        model.ShouldPay = dt.Rows[i][30].ToString().Trim();
                        model.FourInsurance = dt.Rows[i][31].ToString().Trim();
                        model.HealthInsurance = dt.Rows[i][31].ToString().Trim();
                        model.TotalFiveInsurance = dt.Rows[i][33].ToString().Trim();
                        model.ProvidentFundTotal = dt.Rows[i][34].ToString().Trim();
                        model.TotalInsurance = dt.Rows[i][35].ToString().Trim();
                        model.PreTaxSalary = dt.Rows[i][36].ToString().Trim();
                        model.PayableTax = dt.Rows[i][37].ToString().Trim();
                        model.RealWages = dt.Rows[i][38].ToString().Trim();
                        model.JBankCardNumber = dt.Rows[i][39].ToString().Trim();
                        model.Remark = dt.Rows[i][40].ToString().Trim();
                        model.Email = dt.Rows[i][41].ToString().Trim();
                        model.PayDate = dt.Rows[i][42].ToString().Trim();
                        model.PayCycle = dt.Rows[i][43].ToString().Trim();

                        listPayroll10.Add(model);
                    }
                    if (strErrorPayroll10.ToString() != "")
                    {
                        MessageBox.Show(strErrorPayroll10.ToString(), "工资条信息添加错误提示");
                    }
                    else
                    {
                        button1.Invoke(setbutton1enable, new object[] { true });
                    }
                }
                else
                {
                    MessageBox.Show("文件读取为空！");
                }

            }
        }
        #endregion

        #region 发送10号工资
        /// <summary>
        /// 发送10号工资
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {

            InitializeSendPayroll10Timer();
            Thread threadUploadProduct = new Thread(new ThreadStart(Send10));
            threadUploadProduct.SetApartmentState(ApartmentState.STA);
            threadUploadProduct.Start();
        }

        public void Send10()
        {
            button1.Invoke(setbutton1enable, new object[] { false });
            button3.Invoke(setbutton3enable, new object[] { false });

            string fjrtxt = fjr.Text;
            string mmtxt = mm.Text;
            if (fjrtxt == "" || mmtxt == "")
            {
                strErrorPayroll10.Append("亲，请先填写发件邮箱和密码哦！");
                isErrorPayroll10 = true;
            }
            StringBuilder nrStr = new StringBuilder();
            if (listPayroll10 != null && !isErrorPayroll10)
            {
                total10 = listPayroll10.Count;
                current10 = 0;
                for (int i = 0; i < listPayroll10.Count; i++)
                {
                    nrStr = new StringBuilder();
                    nrStr.Append("<h3>美家在线薪资单</h3>");
                    nrStr.Append("您好！感谢您的辛勤劳动和付出，您上月的工资明细请见邮件正文，谢谢！");
                    nrStr.Append("<table style=\"border:1px\">");
                    nrStr.Append("<tr><td>发薪日期</td><td>" + listPayroll10[i].PayDate + "</td></tr>");
                    nrStr.Append("<tr><td>发薪周期</td><td>" + listPayroll10[i].PayCycle + "</td></tr>");
                    nrStr.Append("<tr><td>编号</td><td>" + listPayroll10[i].ID + "</td></tr>");
                    nrStr.Append("<tr><td>姓名</td><td>" + listPayroll10[i].Name + "</td></tr>");
                    nrStr.Append("<tr><td>部门</td><td>" + listPayroll10[i].Department + "</td></tr>");
                    nrStr.Append("<tr><td>职级</td><td>" + listPayroll10[i].Rank + "</td></tr>");
                    nrStr.Append("<tr><td>入职日期</td><td>" + listPayroll10[i].EntryDate + "</td></tr>");
                    nrStr.Append("<tr><td>转正日期</td><td>" + listPayroll10[i].PositiveDate + "</td></tr>");
                    nrStr.Append("<tr><td>当月计薪日</td><td>" + listPayroll10[i].MonthPlan + "</td></tr>");
                    nrStr.Append("<tr><td>转正天数</td><td>" + listPayroll10[i].PositiveDays + "</td></tr>");
                    nrStr.Append("<tr><td>试用薪资</td><td>" + listPayroll10[i].ProbationSalary + "</td></tr>");
                    nrStr.Append("<tr><td>转正薪资</td><td>" + listPayroll10[i].PositiveSalary + "</td></tr>");
                    nrStr.Append("<tr><td>当月薪资标准</td><td>" + listPayroll10[i].MonthSalaryStandard + "</td></tr>");
                    nrStr.Append("<tr><td>月度固定工资</td><td>" + listPayroll10[i].MonthlyFixedSalary + "</td></tr>");
                    nrStr.Append("<tr><td>基本工资（10日)</td><td>" + listPayroll10[i].BasePay + "</td></tr>");
                    nrStr.Append("<tr><td>基本工资补（15日）</td><td>" + listPayroll10[i].BasicSalaryCompensation + "</td></tr>");
                    nrStr.Append("<tr><td>月度绩效薪资基数（15日）</td><td>" + listPayroll10[i].MonthlyPerformanceBaseSalary + "</td></tr>");
                    nrStr.Append("<tr><td>司龄</td><td>" + listPayroll10[i].Seniority + "</td></tr>");
                    nrStr.Append("<tr><td>司龄工资</td><td>" + listPayroll10[i].AgeSalary + "</td></tr>");
                    nrStr.Append("<tr><td>出差补助</td><td>" + listPayroll10[i].TravelAllowance + "</td></tr>");
                    nrStr.Append("<tr><td>全勤奖</td><td>" + listPayroll10[i].FullAttendenceAward + "</td></tr>");
                    nrStr.Append("<tr><td>增加工资合计</td><td>" + listPayroll10[i].IncreaseWagesTotal + "</td></tr>");
                    nrStr.Append("<tr><td>不在岗天数</td><td>" + listPayroll10[i].NotPostDays + "</td></tr>");
                    nrStr.Append("<tr><td>不在岗扣款</td><td>" + listPayroll10[i].NotInCharge + "</td></tr>");
                    nrStr.Append("<tr><td>事假天数</td><td>" + listPayroll10[i].LeaveDays + "</td></tr>");
                    nrStr.Append("<tr><td>事假扣款</td><td>" + listPayroll10[i].LeaveFrom + "</td></tr>");
                    nrStr.Append("<tr><td>病假天数</td><td>" + listPayroll10[i].SickLeaveDays + "</td></tr>");
                    nrStr.Append("<tr><td>病假扣款</td><td>" + listPayroll10[i].SickLeaveDeductions + "</td></tr>");
                    nrStr.Append("<tr><td>迟到/未打卡折算小时数</td><td>" + listPayroll10[i].LateHourNotPunchAt + "</td></tr>");
                    nrStr.Append("<tr><td>迟到扣款</td><td>" + listPayroll10[i].LatePayment + "</td></tr>");
                    nrStr.Append("<tr><td>其他扣款</td><td>" + listPayroll10[i].OtherDeductions + "</td></tr>");
                    nrStr.Append("<tr><td>扣款合计</td><td>" + listPayroll10[i].TotalDeductions + "</td></tr>");
                    nrStr.Append("<tr><td>应发工资</td><td>" + listPayroll10[i].ShouldPay + "</td></tr>");
                    nrStr.Append("<tr><td>四险</td><td>" + listPayroll10[i].FourInsurance + "</td></tr>");
                    nrStr.Append("<tr><td>医保</td><td>" + listPayroll10[i].HealthInsurance + "</td></tr>");
                    nrStr.Append("<tr><td>五险合计</td><td>" + listPayroll10[i].TotalFiveInsurance + "</td></tr>");
                    nrStr.Append("<tr><td>公积金合计</td><td>" + listPayroll10[i].ProvidentFundTotal + "</td></tr>");
                    nrStr.Append("<tr><td>保险合计</td><td>" + listPayroll10[i].TotalInsurance + "</td></tr>");
                    nrStr.Append("<tr><td>税前工资</td><td>" + listPayroll10[i].PreTaxSalary + "</td></tr>");
                    nrStr.Append("<tr><td>应交个税</td><td>" + listPayroll10[i].PayableTax + "</td></tr>");
                    nrStr.Append("<tr><td>实发工资</td><td>" + listPayroll10[i].RealWages + "</td></tr>");
                    nrStr.Append("<tr><td>交行卡号</td><td>" + listPayroll10[i].JBankCardNumber + "</td></tr>");
                    nrStr.Append("<tr><td>备注</td><td>" + listPayroll10[i].Remark + "</td></tr>");
                    nrStr.Append("</table>");
                    string sjrtxt = listPayroll10[i].Email.Trim();//"1181099578@qq.com"; //sjr.Text;
                    string zttxt = listPayroll10[i].Name + "10号工资条";//zt.Text;
                    if (SendEmail(fjrtxt, mmtxt, sjrtxt, zttxt, nrStr.ToString()))
                    {

                        current10++;
                    }
                    else
                    {

                        strErrorPayroll10.Append("第" + (i + 1) + "行工资条发送失败");
                        isErrorPayroll10 = true;
                    }
                }
                if (isErrorPayroll10 == false)
                {
                    MessageBox.Show("10号工资条发送成功");
                    button1.Invoke(setbutton1enable, new object[] { true });
                    button3.Invoke(setbutton3enable, new object[] { true });
                }
                else
                {
                    MessageBox.Show(strErrorPayroll10.ToString() + "\r\n");
                }

            }
            else
            {
                MessageBox.Show(strErrorPayroll10.ToString() + "\r\n");
                strErrorPayroll10 = new StringBuilder();
                button1.Invoke(setbutton1enable, new object[] { true });
                button3.Invoke(setbutton3enable, new object[] { true });
            }
        }

        #endregion


        #region 转换Excel为DataTable
        private DataTable ExcelToDataTable(string strExcelFileName, string strSheetName)
        {
            //通过Jet 引擎或者ACE引擎连接Excel数据源,建议采用ACE引擎，可以支持高版本excel文件
            //string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + strExcelFileName + ";" + "Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + strExcelFileName + ";" + "Extended Properties='Excel 12.0;HDR=NO;IMEX=0';";
            if (strSheetName != "sheet1")
            {
                MessageBox.Show("亲，发送文件命名须为：sheet1");
                return null;
            }
            else
            {
                //Sql语句
                string strExcel = string.Format("select * from [{0}$]", strSheetName); //这是一种方法
                //定义存放的数据表
                DataSet ds = new DataSet();
                //连接数据源

                try
                {
                    OleDbConnection conn = new OleDbConnection(strConn);
                    conn.Open();
                    //适配到数据源
                    OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, strConn);
                    adapter.Fill(ds, strSheetName);

                    conn.Close();

                    return ds.Tables[strSheetName];
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
        }
        #endregion
        #region 发送邮件公共方法
        /// <summary>
        /// 判断是否是正确的Email地址
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        private bool IsEmail(string email)
        {
            Regex rgx = new Regex("(?<user>[^@]+)@(?<host>.+)");

            Match m = rgx.Match(email);

            return m.Success;

        }
        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="fjrtxt"></param>
        /// <param name="mmtxt"></param>
        /// <param name="sjrtxt"></param>
        /// <param name="zttxt"></param>
        /// <param name="nrtxt"></param>
        /// <returns></returns>
        public bool SendEmail(string fjrtxt, string mmtxt, string sjrtxt, string zttxt, string nrtxt)
        {
            if (!IsEmail(fjrtxt))
            {
                return false;
            }
            if (!IsEmail(sjrtxt))
            {
                return false;
            }
            string[] fasong = fjrtxt.Split('@');
            string[] fs = fasong[1].Split('.');
            //发送
            SmtpClient client = new SmtpClient("smtp." + fs[0].ToString().Trim() + ".com");   //设置邮件协议
            client.UseDefaultCredentials = false;//这一句得写前面
            client.DeliveryMethod = SmtpDeliveryMethod.Network; //通过网络发送到Smtp服务器
            client.Credentials = new NetworkCredential(fasong[0].ToString(), mmtxt); //通过用户名和密码 认证
            MailMessage mmsg = new MailMessage(new MailAddress(fjrtxt), new MailAddress(sjrtxt)); //发件人和收件人的邮箱地址
            mmsg.Subject = zttxt;      //邮件主题
            mmsg.SubjectEncoding = Encoding.UTF8;   //主题编码
            mmsg.Body = nrtxt;         //邮件正文
            mmsg.BodyEncoding = Encoding.UTF8;      //正文编码
            mmsg.IsBodyHtml = true;    //设置为HTML格式           
            mmsg.Priority = MailPriority.High;   //优先级
            try
            {
                client.Send(mmsg);
                return true;
                //MessageBox.Show("邮件已发成功");
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return false;
            }
        }
        #endregion


        #region 导入15号工资Excel
        private void button4_Click(object sender, EventArgs e)
        {
            Thread threadSelectExcel15 = new Thread(new ThreadStart(SelectExcel15));
            threadSelectExcel15.SetApartmentState(ApartmentState.STA);
            threadSelectExcel15.Start();
        }
        /// <summary>
        /// 读入15号工资Excel
        /// </summary>
        private void SelectExcel15()
        {
            //strErrorScene.Clear();
            OpenFileDialog PayrollFile15 = new OpenFileDialog();
            PayrollFile15.Title = "15号工资Excel文件上传";
            PayrollFile15.Filter = "Excel文件|*.xls;*.xlsx";//只选择Excel文件
            if (PayrollFile15.ShowDialog() == DialogResult.OK)
            {
                listPayroll15 = new List<Payroll15>();
                txtPayroll15.Invoke(setpayroll15txtfilename, new object[] { PayrollFile15.FileName });
                DataTable dt = ExcelToDataTable(PayrollFile15.FileName, "sheet1");
                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Payroll15 model = new Payroll15();
                        model.ID = dt.Rows[i][0].ToString().Trim();
                        model.Name = dt.Rows[i][1].ToString().Trim();
                        model.Department = dt.Rows[i][2].ToString().Trim();
                        model.MonthPlan = dt.Rows[i][3].ToString().Trim();
                        model.PositiveDays = dt.Rows[i][4].ToString().Trim();
                        model.NotPostDays = dt.Rows[i][5].ToString().Trim();
                        model.ProbationSalary = dt.Rows[i][6].ToString().Trim();
                        model.PositiveSalary = dt.Rows[i][7].ToString().Trim();
                        model.MonthSalaryStandard = dt.Rows[i][8].ToString().Trim();
                        model.MonthlyFixedSalary = dt.Rows[i][9].ToString().Trim();
                        model.MonthlyPerformanceBaseSalary = dt.Rows[i][10].ToString().Trim();
                        model.MonthlyPerformanceCoefficient = dt.Rows[i][11].ToString().Trim();
                        model.MonthlyPerformanceSalary = dt.Rows[i][12].ToString().Trim();
                        model.MonthlyPostPerformanceSalary = dt.Rows[i][13].ToString().Trim();
                        model.BasicSalaryCompensation = dt.Rows[i][14].ToString().Trim();
                        model.NotInCharge = dt.Rows[i][15].ToString().Trim();
                        model.OnBasicSalaryCompensation = dt.Rows[i][16].ToString().Trim();
                        model.PercentageOfWages = dt.Rows[i][17].ToString().Trim();
                        model.ClassAllowance = dt.Rows[i][18].ToString().Trim();
                        model.CompletionCommission = dt.Rows[i][19].ToString().Trim();
                        model.SalesCommissions = dt.Rows[i][20].ToString().Trim();
                        model.ReturnDeduction = dt.Rows[i][21].ToString().Trim();
                        model.Allowance = dt.Rows[i][22].ToString().Trim();
                        model.Oil_TransportationSubsidyStandard = dt.Rows[i][23].ToString().Trim();
                        model.SpecialDuties = dt.Rows[i][24].ToString().Trim();
                        model.OtherSupplement = dt.Rows[i][25].ToString().Trim();
                        model.OtherDeductions = dt.Rows[i][26].ToString().Trim();
                        model.TotalDeductions = dt.Rows[i][27].ToString().Trim();
                        model.TotalWages = dt.Rows[i][28].ToString().Trim();
                        model.JianBankCardNumber = dt.Rows[i][29].ToString().Trim();
                        model.SalaryPaymentDate = dt.Rows[i][30].ToString().Trim();
                        model.Remark = dt.Rows[i][31].ToString().Trim();
                        model.Email = dt.Rows[i][32].ToString().Trim();
                        model.PayCycle = dt.Rows[i][33].ToString().Trim();

                        listPayroll15.Add(model);
                    }

                    if (strErrorPayroll15.ToString() != "")
                    {
                        MessageBox.Show(strErrorPayroll15.ToString(), "工资条信息添加错误提示");
                    }
                    else
                    {
                        button5.Invoke(setbutton5enable, new object[] { true });
                    }
                }
            
            else
            {
                MessageBox.Show("文件读取为空！");
            }
            }
        }

        #endregion

        #region 发送15号工资
        private void button5_Click(object sender, EventArgs e)
        {
            InitializeSendPayroll15Timer();
            Thread threadUploadProduct = new Thread(new ThreadStart(Send15));
            threadUploadProduct.SetApartmentState(ApartmentState.STA);
            threadUploadProduct.Start();
        }
        public void Send15()
        {
            button4.Invoke(setbutton4enable, new object[] { false });
            button5.Invoke(setbutton5enable, new object[] { false });

            string fjrtxt = fjr.Text;
            string mmtxt = mm.Text;
            if (fjrtxt == "" || mmtxt == "")
            {
                strErrorPayroll15.Append("亲，请先填写发件邮箱和密码哦！");
                isErrorPayroll15 = true;
            }
            StringBuilder nrStr = new StringBuilder();
            if (listPayroll15 != null && !isErrorPayroll15)
            {
                total15 = listPayroll15.Count;
                current15 = 0;
                for (int i = 0; i < listPayroll15.Count; i++)
                {
                    nrStr = new StringBuilder();
                    nrStr.Append("<h3>美家在线薪资单</h3>");
                    nrStr.Append("您好！感谢您的辛勤劳动和付出，您上月的工资明细请见邮件正文，谢谢！");
                    nrStr.Append("<table>");
                    nrStr.Append("<tr><td>发薪日期</td><td>" + listPayroll15[i].SalaryPaymentDate + "</td></tr>");
                    nrStr.Append("<tr><td>发薪周期</td><td>" + listPayroll15[i].PayCycle + "</td></tr>");
                    nrStr.Append("<tr><td>编号</td><td>" + listPayroll15[i].ID + "</td></tr>");
                    nrStr.Append("<tr><td>姓名</td><td  >" + listPayroll15[i].Name + "</td></tr>");
                    nrStr.Append("<tr><td>部门</td><td  >" + listPayroll15[i].Department + "</td></tr>");
                    nrStr.Append("<tr><td>计薪日</td><td  >" + listPayroll15[i].MonthPlan + "</td></tr>");
                    nrStr.Append("<tr><td>转正天数</td><td  >" + listPayroll15[i].PositiveDays + "</td></tr>");
                    nrStr.Append("<tr><td>不在岗天数</td><td  >" + listPayroll15[i].NotPostDays + "</td></tr>");
                    nrStr.Append("<tr><td>试用薪资</td><td  >" + listPayroll15[i].ProbationSalary + "</td></tr>");
                    nrStr.Append("<tr><td>转正薪资</td><td  >" + listPayroll15[i].PositiveSalary + "</td></tr>");
                    nrStr.Append("<tr><td>当月薪资标准</td><td  >" + listPayroll15[i].MonthSalaryStandard + "</td></tr>");
                    nrStr.Append("<tr><td>月度固定工资</td><td  >" + listPayroll15[i].MonthlyFixedSalary + "</td></tr>");
                    nrStr.Append("<tr><td>月度绩效薪资基数</td><td  >" + listPayroll15[i].MonthlyPerformanceBaseSalary + "</td></tr>");
                    nrStr.Append("<tr><td>月度绩效系数</td><td  >" + listPayroll15[i].MonthlyPerformanceCoefficient + "</td></tr>");
                    nrStr.Append("<tr><td>月度绩效工资</td><td  >" + listPayroll15[i].MonthlyPerformanceSalary + "</td></tr>");
                    nrStr.Append("<tr><td>月度在岗绩效工资</td><td  >" + listPayroll15[i].MonthlyPostPerformanceSalary + "</td></tr>");
                    nrStr.Append("<tr><td>基本工资补</td><td  >" + listPayroll15[i].BasicSalaryCompensation + "</td></tr>");
                    nrStr.Append("<tr><td>不在岗扣款</td><td  >" + listPayroll15[i].NotInCharge + "</td></tr>");
                    nrStr.Append("<tr><td>在岗基本工资补</td><td  >" + listPayroll15[i].OnBasicSalaryCompensation + "</td></tr>");
                    nrStr.Append("<tr><td>提成工资</td><td  >" + listPayroll15[i].PercentageOfWages + "</td></tr>");
                    nrStr.Append("<tr><td>课时补贴</td><td  >" + listPayroll15[i].ClassAllowance + "</td></tr>");
                    nrStr.Append("<tr><td>竣工提成</td><td  >" + listPayroll15[i].CompletionCommission + "</td></tr>");
                    nrStr.Append("<tr><td>销售提成</td><td　 >" + listPayroll15[i].SalesCommissions + "</td></tr>");
                    nrStr.Append("<tr><td>退货扣减</td><td　 >" + listPayroll15[i].ReturnDeduction + "</td></tr>");
                    nrStr.Append("<tr><td>补贴款</td><td　 >" + listPayroll15[i].Allowance + "</td></tr>");
                    nrStr.Append("<tr><td>油补/交通补助标准</td><td  >" + listPayroll15[i].Oil_TransportationSubsidyStandard + "</td></tr>");
                    nrStr.Append("<tr><td>加班餐费</td><td  >" + listPayroll15[i].SpecialDuties + "</td></tr>");
                    nrStr.Append("<tr><td>其他补款</td><td  >" + listPayroll15[i].OtherSupplement + "</td></tr>");
                    nrStr.Append("<tr><td>其他扣款</td><td  >" + listPayroll15[i].OtherDeductions + "</td></tr>");
                    nrStr.Append("<tr><td>扣款合计</td><td  >" + listPayroll15[i].TotalDeductions + "</td></tr>");
                    nrStr.Append("<tr><td>工资合计</td><td  >" + listPayroll15[i].TotalWages + "</td></tr>");
                    nrStr.Append("<tr><td>建行卡号</td><td  >" + listPayroll15[i].JianBankCardNumber + "</td></tr>");
                    nrStr.Append("<tr><td>备注</td><td  >" + listPayroll15[i].Remark + "</td></tr>");


                    nrStr.Append("</table>");
                    string sjrtxt = listPayroll15[i].Email.Trim();//"1181099578@qq.com"; //sjr.Text;
                    string zttxt = listPayroll15[i].Name + "15号工资条";//zt.Text;
                    if (SendEmail(fjrtxt, mmtxt, sjrtxt, zttxt, nrStr.ToString()))
                    {
                        current15++;
                    }
                    else
                    {
                        strErrorPayroll15.Append("第" + (i + 1) + "行工资条发送失败");
                        isErrorPayroll15 = true;
                    }
                }
                if (isErrorPayroll15 == false)
                {
                    MessageBox.Show("15号工资条发送成功");
                    button4.Invoke(setbutton4enable, new object[] { true });
                    button5.Invoke(setbutton5enable, new object[] { true });
                }
                else
                {
                    MessageBox.Show(strErrorPayroll15.ToString() + "\r\n");
                }

            }
            else
            {
                MessageBox.Show(strErrorPayroll15.ToString() + "\r\n");
                strErrorPayroll15 = new StringBuilder();
                button4.Invoke(setbutton4enable, new object[] { true });
                button5.Invoke(setbutton5enable, new object[] { true });
            }
        }

        #endregion



        #region 导入15号管理者工资Excel
        private void button6_Click(object sender, EventArgs e)
        {
            Thread threadSelectExcelMa15 = new Thread(new ThreadStart(SelectExcelMa15));
            threadSelectExcelMa15.SetApartmentState(ApartmentState.STA);
            threadSelectExcelMa15.Start();
        }
        /// <summary>
        /// 读入15号管理者工资Excel
        /// </summary>
        private void SelectExcelMa15()
        {
            //strErrorScene.Clear();
            OpenFileDialog MaPayrollFile15 = new OpenFileDialog();
            MaPayrollFile15.Title = "15号管理者工资Excel文件上传";
            MaPayrollFile15.Filter = "Excel文件|*.xls;*.xlsx";//只选择Excel文件
            if (MaPayrollFile15.ShowDialog() == DialogResult.OK)
            {
                listMaPayroll15 = new List<MaPayroll15>();
                txtMaPayroll15.Invoke(setmapayroll15txtfilename, new object[] { MaPayrollFile15.FileName });
                DataTable dt = ExcelToDataTable(MaPayrollFile15.FileName, "sheet1");
                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        MaPayroll15 model = new MaPayroll15();
                        model.ID = dt.Rows[i][0].ToString().Trim();
                        model.Name = dt.Rows[i][1].ToString().Trim();
                        model.Department = dt.Rows[i][2].ToString().Trim();
                        model.MonthPlan = dt.Rows[i][3].ToString().Trim();
                        model.PositiveDays = dt.Rows[i][4].ToString().Trim();
                        model.NotPostDays = dt.Rows[i][5].ToString().Trim();
                        model.ProbationSalary = dt.Rows[i][6].ToString().Trim();
                        model.PositiveSalary = dt.Rows[i][7].ToString().Trim();
                        model.MonthSalaryStandard = dt.Rows[i][8].ToString().Trim();
                        model.MonthlyFixedSalary = dt.Rows[i][9].ToString().Trim();
                        model.QuarterlyPerformanceConefficient = dt.Rows[i][10].ToString().Trim();
                        model.TimeFactor = dt.Rows[i][11].ToString().Trim();
                        model.QuarterlyPerformanceSalary = dt.Rows[i][12].ToString().Trim();
                        model.BasicSalaryCompensation = dt.Rows[i][13].ToString().Trim();
                        model.NotInCharge = dt.Rows[i][14].ToString().Trim();
                        model.OnBasicSalaryCompensation = dt.Rows[i][15].ToString().Trim();
                        model.PercentageOfWages = dt.Rows[i][16].ToString().Trim();
                        model.ClassAllowance = dt.Rows[i][17].ToString().Trim();
                        model.CompletionCommission = dt.Rows[i][18].ToString().Trim();
                        model.SalesCommissions = dt.Rows[i][19].ToString().Trim();
                        model.ReturnDeduction = dt.Rows[i][20].ToString().Trim();
                        model.Allowance = dt.Rows[i][21].ToString().Trim();
                        model.Oil_TransportationSubsidyStandard = dt.Rows[i][22].ToString().Trim();
                        model.SpecialDuties = dt.Rows[i][23].ToString().Trim();
                        model.OtherSupplement = dt.Rows[i][24].ToString().Trim();
                        model.OtherDeductions = dt.Rows[i][25].ToString().Trim();
                        model.TotalDeductions = dt.Rows[i][26].ToString().Trim();
                        model.TotalWages = dt.Rows[i][27].ToString().Trim();
                        model.JianBankCardNumber = dt.Rows[i][28].ToString().Trim();
                        model.SalaryPaymentDate = dt.Rows[i][29].ToString().Trim();
                        model.PayCycle = dt.Rows[i][30].ToString().Trim();
                        model.Remark = dt.Rows[i][31].ToString().Trim();
                        model.Email = dt.Rows[i][32].ToString().Trim();


                        listMaPayroll15.Add(model);

                    }
                    if (strErrorMaPayroll15.ToString() != "")
                    {
                        MessageBox.Show(strErrorMaPayroll15.ToString(), "工资条信息添加错误提示");
                    }
                    else
                    {
                        button7.Invoke(setbutton7enable, new object[] { true });
                    }
                }
                else
                {
                    MessageBox.Show("文件读取为空！");
                }
            }
        }

        #endregion

        #region 发送15号工资
        private void button7_Click(object sender, EventArgs e)
        {
            InitializeSendPayroll15Timer();
            Thread threadUploadProduct = new Thread(new ThreadStart(SendMa15));
            threadUploadProduct.SetApartmentState(ApartmentState.STA);
            threadUploadProduct.Start();
        }
        public void SendMa15()
        {
            button6.Invoke(setbutton6enable, new object[] { false });
            button7.Invoke(setbutton7enable, new object[] { false });

            string fjrtxt = fjr.Text;
            string mmtxt = mm.Text;
            if (fjrtxt == "" || mmtxt == "")
            {
                strErrorMaPayroll15.Append("亲，请先填写发件邮箱和密码哦！");
                isErrorMaPayroll15 = true;
            }
            StringBuilder nrStr = new StringBuilder();
            if (listMaPayroll15 != null && !isErrorMaPayroll15)
            {
                totalMa15 = listMaPayroll15.Count;
                currentMa15 = 0;
                for (int i = 0; i < listMaPayroll15.Count; i++)
                {
                    nrStr = new StringBuilder();
                    nrStr.Append("<h3>美家在线薪资单</h3>");
                    nrStr.Append("您好！感谢您的辛勤劳动和付出，您上月的工资明细请见邮件正文，谢谢！");
                    nrStr.Append("<table>");
                    nrStr.Append("<tr><td>发薪日期</td><td>" + listMaPayroll15[i].SalaryPaymentDate + "</td></tr>");
                    nrStr.Append("<tr><td>发薪周期</td><td>" + listMaPayroll15[i].PayCycle + "</td></tr>");
                    nrStr.Append("<tr><td>编号</td><td>" + listMaPayroll15[i].ID + "</td></tr>");
                    nrStr.Append("<tr><td>姓名</td><td  >" + listMaPayroll15[i].Name + "</td></tr>");
                    nrStr.Append("<tr><td>部门</td><td  >" + listMaPayroll15[i].Department + "</td></tr>");
                    nrStr.Append("<tr><td>计薪日</td><td  >" + listMaPayroll15[i].MonthPlan + "</td></tr>");
                    nrStr.Append("<tr><td>转正天数</td><td  >" + listMaPayroll15[i].PositiveDays + "</td></tr>");
                    nrStr.Append("<tr><td>不在岗天数</td><td  >" + listMaPayroll15[i].NotPostDays + "</td></tr>");
                    nrStr.Append("<tr><td>试用薪资</td><td  >" + listMaPayroll15[i].ProbationSalary + "</td></tr>");
                    nrStr.Append("<tr><td>转正薪资</td><td  >" + listMaPayroll15[i].PositiveSalary + "</td></tr>");
                    nrStr.Append("<tr><td>当月薪资标准</td><td  >" + listMaPayroll15[i].MonthSalaryStandard + "</td></tr>");
                    nrStr.Append("<tr><td>月度固定工资</td><td  >" + listMaPayroll15[i].MonthlyFixedSalary + "</td></tr>");
                    nrStr.Append("<tr><td>季度绩效系数</td><td  >" + listMaPayroll15[i].QuarterlyPerformanceConefficient + "</td></tr>");
                    nrStr.Append("<tr><td>时间系数</td><td  >" + listMaPayroll15[i].TimeFactor + "</td></tr>");
                    nrStr.Append("<tr><td>季度绩效工资</td><td  >" + listMaPayroll15[i].QuarterlyPerformanceSalary + "</td></tr>");
                    nrStr.Append("<tr><td>基本工资补</td><td  >" + listMaPayroll15[i].BasicSalaryCompensation + "</td></tr>");
                    nrStr.Append("<tr><td>不在岗扣款</td><td  >" + listMaPayroll15[i].NotInCharge + "</td></tr>");
                    nrStr.Append("<tr><td>在岗基本工资补</td><td  >" + listMaPayroll15[i].OnBasicSalaryCompensation + "</td></tr>");
                    nrStr.Append("<tr><td>提成工资</td><td  >" + listMaPayroll15[i].PercentageOfWages + "</td></tr>");
                    nrStr.Append("<tr><td>课时补贴</td><td  >" + listMaPayroll15[i].ClassAllowance + "</td></tr>");
                    nrStr.Append("<tr><td>竣工提成</td><td  >" + listMaPayroll15[i].CompletionCommission + "</td></tr>");
                    nrStr.Append("<tr><td>销售提成</td><td　 >" + listMaPayroll15[i].SalesCommissions + "</td></tr>");
                    nrStr.Append("<tr><td>退货扣减</td><td　 >" + listMaPayroll15[i].ReturnDeduction + "</td></tr>");
                    nrStr.Append("<tr><td>补贴款</td><td　 >" + listMaPayroll15[i].Allowance + "</td></tr>");
                    nrStr.Append("<tr><td>油补/交通补助标准</td><td  >" + listMaPayroll15[i].Oil_TransportationSubsidyStandard + "</td></tr>");
                    nrStr.Append("<tr><td>加班餐费</td><td  >" + listMaPayroll15[i].SpecialDuties + "</td></tr>");
                    nrStr.Append("<tr><td>其他补款</td><td  >" + listMaPayroll15[i].OtherSupplement + "</td></tr>");
                    nrStr.Append("<tr><td>其他扣款</td><td  >" + listMaPayroll15[i].OtherDeductions + "</td></tr>");
                    nrStr.Append("<tr><td>扣款合计</td><td  >" + listMaPayroll15[i].TotalDeductions + "</td></tr>");
                    nrStr.Append("<tr><td>工资合计</td><td  >" + listMaPayroll15[i].TotalWages + "</td></tr>");
                    nrStr.Append("<tr><td>建行卡号</td><td  >" + listMaPayroll15[i].JianBankCardNumber + "</td></tr>");
                    nrStr.Append("<tr><td>备注</td><td  >" + listMaPayroll15[i].Remark + "</td></tr>");


                    nrStr.Append("</table>");
                    string sjrtxt = listMaPayroll15[i].Email.Trim();//"1181099578@qq.com"; //sjr.Text;
                    string zttxt = listMaPayroll15[i].Name + "15号管理者工资条";//zt.Text;
                    if (SendEmail(fjrtxt, mmtxt, sjrtxt, zttxt, nrStr.ToString()))
                    {
                        currentMa15++;
                    }
                    else
                    {
                        strErrorMaPayroll15.Append("第" + (i + 1) + "行工资条发送失败");
                        isErrorMaPayroll15 = true;
                    }
                }
                if (isErrorMaPayroll15 == false)
                {
                    MessageBox.Show("15号管理者工资条发送成功");
                    button6.Invoke(setbutton6enable, new object[] { true });
                    button7.Invoke(setbutton7enable, new object[] { true });
                }
                else
                {
                    MessageBox.Show(strErrorMaPayroll15.ToString() + "\r\n");
                }

            }
            else
            {
                MessageBox.Show(strErrorMaPayroll15.ToString() + "\r\n");
                strErrorMaPayroll15 = new StringBuilder();
                button6.Invoke(setbutton6enable, new object[] { true });
                button7.Invoke(setbutton7enable, new object[] { true });
            }
        }

        #endregion
        #region 进度条
        System.Windows.Forms.Timer sendPayroll10Timer = new System.Windows.Forms.Timer();
        System.Windows.Forms.Timer sendPayroll15Timer = new System.Windows.Forms.Timer();
        System.Windows.Forms.Timer sendMaPayroll15Timer = new System.Windows.Forms.Timer();
        System.Windows.Forms.Timer sendMaPayrollYearAwardTimer = new System.Windows.Forms.Timer();

        //10号
        private void InitializeSendPayroll10Timer()
        {
            sendPayroll10Timer.Interval = 100;
            sendPayroll10Timer.Tick += new EventHandler(IncreaseSendPayroll10ProgressBar);
            sendPayroll10Timer.Start();
        }

        private void IncreaseSendPayroll10ProgressBar(object sender, EventArgs e)
        {

            try
            {
                progressBar1.Maximum = 100;
                progressBar1.Value = ((current10 * 100) / total10);
                if (progressBar1.Value >= progressBar1.Maximum)
                {
                    sendPayroll10Timer.Stop();
                }
            }
            catch
            {
                sendPayroll10Timer.Stop();
            }

        }

        //15号员工
        private void InitializeSendPayroll15Timer()
        {
            sendPayroll15Timer.Interval = 100;
            sendPayroll15Timer.Tick += new EventHandler(IncreaseSendPayroll15ProgressBar);
            sendPayroll15Timer.Start();
        }

        private void IncreaseSendPayroll15ProgressBar(object sender, EventArgs e)
        {

            try
            {
                progressBar2.Maximum = 100;
                progressBar2.Value = ((current15 * 100) / total15);
                if (progressBar2.Value >= progressBar2.Maximum)
                {
                    sendPayroll15Timer.Stop();
                }
            }
            catch
            {
                sendPayroll15Timer.Stop();
            }

        }

        //15号管理者
        private void InitializeSendMaPayroll15Timer()
        {
            sendMaPayroll15Timer.Interval = 100;
            sendMaPayroll15Timer.Tick += new EventHandler(IncreaseSendMaPayroll15ProgressBar);
            sendMaPayroll15Timer.Start();
        }

        private void IncreaseSendMaPayroll15ProgressBar(object sender, EventArgs e)
        {

            try
            {
                progressBar3.Maximum = 100;
                progressBar3.Value = ((currentMa15 * 100) / totalMa15);
                if (progressBar3.Value >= progressBar3.Maximum)
                {
                    sendMaPayroll15Timer.Stop();
                }
            }
            catch
            {
                sendMaPayroll15Timer.Stop();
            }

        }

        //年终奖
        private void InitializeSendPayrollYearAwardTimer()
        {
            sendMaPayrollYearAwardTimer.Interval = 100;
            sendMaPayrollYearAwardTimer.Tick += new EventHandler(IncreaseSendYearAwardProgressBar);
            sendMaPayrollYearAwardTimer.Start();
        }

        private void IncreaseSendYearAwardProgressBar(object sender, EventArgs e)
        {

            try
            {
                progressBar4.Maximum = 100;
                progressBar4.Value = ((currentYearAward * 100) / totalYearAward);
                if (progressBar3.Value >= progressBar4.Maximum)
                {
                    sendMaPayrollYearAwardTimer.Stop();
                }
            }
            catch
            {
                sendMaPayrollYearAwardTimer.Stop();
            }

        }

       

        #endregion
        #region 导入年终奖工资Excel
        private void button9_Click(object sender, EventArgs e)
        {
            Thread threadSelectExcelYearAward = new Thread(new ThreadStart(SelectExcelYearAward));
            threadSelectExcelYearAward.SetApartmentState(ApartmentState.STA);
            threadSelectExcelYearAward.Start();
        }


        /// <summary>
        /// 读入年终奖Excel
        /// </summary>
        private void SelectExcelYearAward()
        {
            OpenFileDialog PayrollFileYearAward = new OpenFileDialog();
            PayrollFileYearAward.Title = "年终奖Excel文件上传";
            PayrollFileYearAward.Filter = "Excel文件|*.xls;*.xlsx";//只选择Excel文件
            if (PayrollFileYearAward.ShowDialog() == DialogResult.OK)
            {
                listYearAward = new List<YearAwardPayroll>();
                txtYearAward.Invoke(setyearawardtxtfilename, new object[] { PayrollFileYearAward.FileName });
                DataTable dt = ExcelToDataTable(PayrollFileYearAward.FileName, "sheet1");
                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        YearAwardPayroll model = new YearAwardPayroll();
                        model.ID = dt.Rows[i][0].ToString().Trim();
                        model.Name = dt.Rows[i][1].ToString().Trim();
                        model.Department = dt.Rows[i][2].ToString().Trim();
                        model.EntryTime = dt.Rows[i][3].ToString().Trim();
                        model.Deadline = dt.Rows[i][4].ToString().Trim();
                        model.TimeFactor = dt.Rows[i][5].ToString().Trim();
                        model.WageStandard = string.Format("{0:0000.00}",dt.Rows[i][6].ToString().Trim());
                        model.AvePerCoefficient = dt.Rows[i][7].ToString().Trim();
                        model.YearAward = string.Format("{0:0000.00}", dt.Rows[i][8].ToString().Trim());
                        model.OtherIncentives = string.Format("{0:0000.00}", dt.Rows[i][9].ToString().Trim());
                        model.YearEndBonus = string.Format("{0:0000.00}", dt.Rows[i][10].ToString().Trim());
                        model.JianBankCardNumber = dt.Rows[i][11].ToString().Trim();
                        model.SalaryPaymentDate = dt.Rows[i][12].ToString().Trim(); 
                        model.PayCycle = dt.Rows[i][13].ToString().Trim();
                        model.Email = dt.Rows[i][14].ToString().Trim();
                        model.Remark = dt.Rows[i][15].ToString().Trim();

                        listYearAward.Add(model);
                    }

                    if (strErrorYearAwardPayroll.ToString() != "")
                    {
                        MessageBox.Show(strErrorYearAwardPayroll.ToString(), "工资条信息添加错误提示");
                    }
                    else
                    {
                        button8.Invoke(setbutton8enable, new object[] { true });
                    }
                }

                else
                {
                    MessageBox.Show("文件读取为空！");
                }
            }
        }
        #endregion

        #region 发送年终奖
        private void button8_Click(object sender, EventArgs e)
        {
            InitializeSendPayrollYearAwardTimer();
            Thread threadUploadYearAward = new Thread(new ThreadStart(SendYearAward));
            threadUploadYearAward.SetApartmentState(ApartmentState.STA);
            threadUploadYearAward.Start();
        }

        public void SendYearAward()
        {
            button8.Invoke(setbutton8enable, new object[] { false });
            button9.Invoke(setbutton9enable, new object[] { false });

            string fjrtxt = fjr.Text;
            string mmtxt = mm.Text;
            if (fjrtxt == "" || mmtxt == "")
            {
                strErrorYearAwardPayroll.Append("亲，请先填写发件邮箱和密码哦！");
                isErrorMaPayroll15 = true;
            }
            StringBuilder nrStr = new StringBuilder();
            if (listYearAward != null && !isErrorYearAwardPayroll)
            {
                totalYearAward = listYearAward.Count;
                currentYearAward = 0;
                for (int i = 0; i < listYearAward.Count; i++)
                {
                    nrStr = new StringBuilder();
                    nrStr.Append("<h3>美家在线薪资单</h3>");
                    string year = DateTime.Now.ToString("yyyy");
                    nrStr.Append("您好！感谢您的辛勤劳动和付出，您的"+year+"年年终奖明细请见邮件正文，谢谢！");
                    nrStr.Append("<table>");
                    nrStr.Append("<tr><td>发薪日期</td><td>" + listYearAward[i].SalaryPaymentDate + "</td></tr>");
                    nrStr.Append("<tr><td>奖金周期</td><td>" + listYearAward[i].PayCycle + "</td></tr>");
                    nrStr.Append("<tr><td>编号</td><td>" + listYearAward[i].ID + "</td></tr>");
                    nrStr.Append("<tr><td>姓名</td><td  >" + listYearAward[i].Name + "</td></tr>");
                    nrStr.Append("<tr><td>部门</td><td  >" + listYearAward[i].Department + "</td></tr>");
                    nrStr.Append("<tr><td>入职时间</td><td  >" + listYearAward[i].EntryTime + "</td></tr>");
                    nrStr.Append("<tr><td>截止日期</td><td  >" + listYearAward[i].Deadline + "</td></tr>");
                    nrStr.Append("<tr><td>时间系数</td><td  >" + listYearAward[i].TimeFactor + "</td></tr>");
                    nrStr.Append("<tr><td>2015平均工资标准</td><td  >" + listYearAward[i].WageStandard + "</td></tr>");
                    nrStr.Append("<tr><td>2015平均绩效系数 </td><td  >" + listYearAward[i].AvePerCoefficient + "</td></tr>");
                    nrStr.Append("<tr><td>年终奖</td><td  >" + listYearAward[i].YearAward + "</td></tr>");
                    nrStr.Append("<tr><td>其他激励</td><td  >" + listYearAward[i].OtherIncentives + "</td></tr>");
                    nrStr.Append("<tr><td>年终奖金合计</td><td  >" + listYearAward[i].YearEndBonus + "</td></tr>");        
                    nrStr.Append("<tr><td>建行卡号</td><td  >" + listYearAward[i].JianBankCardNumber + "</td></tr>");
                    nrStr.Append("<tr><td>备注</td><td  >" + listYearAward[i].Remark + "</td></tr>");
                    nrStr.Append("</table>");
                    string sjrtxt = listYearAward[i].Email.Trim();
                    string zttxt = listYearAward[i].Name + year + "年年终奖工资条";
                    if (SendEmail(fjrtxt, mmtxt, sjrtxt, zttxt, nrStr.ToString()))
                    {
                        currentYearAward++;
                    }
                    else
                    {
                        strErrorYearAwardPayroll.Append("第" + (i + 1) + "行工资条发送失败");
                        isErrorYearAwardPayroll = true;
                    }
                }
                if (isErrorYearAwardPayroll == false)
                {
                    MessageBox.Show("年终奖工资条发送成功");
                    button8.Invoke(setbutton8enable, new object[] { true });
                    button9.Invoke(setbutton9enable, new object[] { true });
                }
                else
                {
                    MessageBox.Show(strErrorYearAwardPayroll.ToString() + "\r\n");
                }

            }
            else
            {
                MessageBox.Show(strErrorYearAwardPayroll.ToString() + "\r\n");
                strErrorYearAwardPayroll = new StringBuilder();
                button8.Invoke(setbutton8enable, new object[] { true });
                button9.Invoke(setbutton9enable, new object[] { true });
            }
        }
        #endregion
       

        #region 退出程序
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

    }


    #region 公共MODEL
    /// <summary>
    /// 10号工资表Model
    /// </summary>
    public class Payroll10
    {
        public string ID { get; set; }//编号
        public string Name { get; set; }//姓名
        public string Department { get; set; }//部门
        public string Rank { get; set; }//职级
        public string EntryDate { get; set; }//入职日期
        public string PositiveDate { get; set; }//转正日期
        public string MonthPlan { get; set; }//当月计薪日
        public string PositiveDays { get; set; }//转正天数
        public string ProbationSalary { get; set; }//试用薪资
        public string PositiveSalary { get; set; }//转正薪资
        public string MonthSalaryStandard { get; set; }//当月薪资标准
        public string MonthlyFixedSalary { get; set; }//月度固定工资
        public string BasePay { get; set; }//基本工资（10日)
        public string BasicSalaryCompensation { get; set; }//基本工资补（15日）
        public string MonthlyPerformanceBaseSalary { get; set; }//月度绩效薪资基数（15日）
        public string Seniority { get; set; }//司龄
        public string AgeSalary { get; set; }//司龄工资
        public string TravelAllowance { get; set; }//出差补助
        public string FullAttendenceAward { get; set; }//全勤奖
        public string IncreaseWagesTotal { get; set; }//增加工资合计
        public string NotPostDays { get; set; }//不在岗天数
        public string NotInCharge { get; set; }//不在岗扣款
        public string LeaveDays { get; set; }//事假天数
        public string LeaveFrom { get; set; }//事假扣款
        public string SickLeaveDays { get; set; }//病假天数
        public string SickLeaveDeductions { get; set; }//病假扣款
        public string LateHourNotPunchAt { get; set; }//迟到/未打卡折算小时数
        public string LatePayment { get; set; }//迟到扣款
        public string OtherDeductions { get; set; }//其他扣款
        public string TotalDeductions { get; set; }//扣款合计
        public string ShouldPay { get; set; }//应发工资
        public string FourInsurance { get; set; }//四险
        public string HealthInsurance { get; set; }//医保
        public string TotalFiveInsurance { get; set; }//五险合计
        public string ProvidentFundTotal { get; set; }//公积金合计
        public string TotalInsurance { get; set; }//保险合计
        public string PreTaxSalary { get; set; }//税前工资
        public string PayableTax { get; set; }//应交个税
        public string RealWages { get; set; }//实发工资
        public string JBankCardNumber { get; set; }//交行卡号
        public string Remark { get; set; }//备注
        public string Email { get; set; }//邮箱
        public string PayDate { get; set; }//发薪日期 
        public string PayCycle { get; set; }//发薪周期
    }

    /// <summary>
    /// 15号员工工资表Model
    /// </summary>
    public class Payroll15
    {
        public string ID { get; set; }//编号
        public string Name { get; set; }//姓名
        public string Department { get; set; }//部门
        public string MonthPlan { get; set; }//计薪日
        public string PositiveDays { get; set; }//转正天数
        public string NotPostDays { get; set; }//不在岗天数
        public string ProbationSalary { get; set; }//试用薪资
        public string PositiveSalary { get; set; }//转正薪资
        public string MonthSalaryStandard { get; set; }//当月薪资标准
        public string MonthlyFixedSalary { get; set; }//月度固定工资
        public string MonthlyPerformanceBaseSalary { get; set; }//月度绩效薪资基数
        public string MonthlyPerformanceCoefficient { get; set; }//月度绩效系数
        public string MonthlyPerformanceSalary { get; set; }//月度绩效工资
        public string MonthlyPostPerformanceSalary { get; set; }//月度在岗绩效工资
        public string BasicSalaryCompensation { get; set; }//基本工资补       
        public string NotInCharge { get; set; }//不在岗扣款
        public string OnBasicSalaryCompensation { get; set; }//在岗基本工资补
        public string PercentageOfWages { get; set; }//提成工资
        public string ClassAllowance { get; set; }//课时补贴
        public string CompletionCommission { get; set; }//竣工提成
        public string SalesCommissions { get; set; }//销售提成
        public string ReturnDeduction { get; set; }//退货扣减
        public string Allowance { get; set; }//补贴款
        public string Oil_TransportationSubsidyStandard { get; set; }//油补/交通补助标准
        public string SpecialDuties { get; set; }//加班餐费
        public string OtherSupplement { get; set; }//其他补款
        public string OtherDeductions { get; set; }//其他扣款
        public string TotalDeductions { get; set; }//扣款合计
        public string TotalWages { get; set; }//工资合计
        public string JianBankCardNumber { get; set; }//建行卡号
        public string SalaryPaymentDate { get; set; }//工资发放日期      
        public string Remark { get; set; }//备注
        public string Email { get; set; }//邮箱
        public string PayCycle { get; set; }//发薪周期
    }

    /// <summary>
    /// 15号管理者工资表Model
    /// </summary>
    public class MaPayroll15
    {
        public string ID { get; set; }//编号
        public string Name { get; set; }//姓名
        public string Department { get; set; }//部门
        public string MonthPlan { get; set; }//计薪日
        public string PositiveDays { get; set; }//转正天数
        public string NotPostDays { get; set; }//不在岗天数
        public string ProbationSalary { get; set; }//试用薪资
        public string PositiveSalary { get; set; }//转正薪资
        public string MonthSalaryStandard { get; set; }//当月薪资标准
        public string MonthlyFixedSalary { get; set; }//月度固定工资
        public string QuarterlyPerformanceConefficient { get; set; }//季度绩效系数
        public string TimeFactor { get; set; }//时间系数
        public string QuarterlyPerformanceSalary { get; set; }//季度绩效工资
        public string BasicSalaryCompensation { get; set; }//基本工资补       
        public string NotInCharge { get; set; }//不在岗扣款
        public string OnBasicSalaryCompensation { get; set; }//在岗基本工资补
        public string PercentageOfWages { get; set; }//提成工资
        public string ClassAllowance { get; set; }//课时补贴
        public string CompletionCommission { get; set; }//竣工提成
        public string SalesCommissions { get; set; }//销售提成
        public string ReturnDeduction { get; set; }//退货扣减
        public string Allowance { get; set; }//补贴款
        public string Oil_TransportationSubsidyStandard { get; set; }//油补/交通补助标准
        public string SpecialDuties { get; set; }//加班餐费
        public string OtherSupplement { get; set; }//其他补款
        public string OtherDeductions { get; set; }//其他扣款
        public string TotalDeductions { get; set; }//扣款合计
        public string TotalWages { get; set; }//工资合计
        public string JianBankCardNumber { get; set; }//建行卡号
        public string SalaryPaymentDate { get; set; }//工资发放日期      
        public string PayCycle { get; set; }//发薪周期
        public string Remark { get; set; }//备注
        public string Email { get; set; }//邮箱


    }


    /// <summary>
    /// 年终奖工资表Model
    /// </summary>
    public class YearAwardPayroll
    {
        public string ID { get; set; }//编号
        public string Name { get; set; }//姓名
        public string Department { get; set; }//部门
        public string EntryTime  { get; set; }//入职时间
        public string Deadline { get; set; }//截止日期
        public string TimeFactor { get; set; }//时间系数
        public string WageStandard { get; set; }//2015平均工资标准
        public string AvePerCoefficient  { get; set; }//2015平均绩效系数       
        public string YearAward { get; set; }//年终奖
        public string OtherIncentives  { get; set; }//其他激励
        public string YearEndBonus  { get; set; }//年终奖金合计
        public string JianBankCardNumber { get; set; }//建行卡号
        public string SalaryPaymentDate { get; set; }//工资发放日期      
        public string PayCycle { get; set; }//奖金周期
        public string Email { get; set; }//邮箱
        public string Remark { get; set; }//备注
    }
    #endregion

}