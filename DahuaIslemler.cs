using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using NetSDKCS;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading;
using Newtonsoft.Json;
using System.Configuration;
using System.Data.SqlClient;
using System.Web.Security;

namespace CSB
{

    public class DahuaIslemler
    {
        public class Constants
        {
            public const int SELECTED_DOOR_INDEX = 0;
            public const int CARD_STATE_NORMAL = 0;
            public const int CARD_TYPE_GENERAL = 0;
            public const int CARD_TYPE_GUEST = 2;
            public const int CARD_TYPE_MOTHERCARD = 0xFF;
            public const bool CARD_FIRST_ENTER_STATE = false;
        }

        public class CardInfo
        {
            public NET_TIME stuCreateTime;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szCardNo;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szUserID;

            public EM_ACCESSCTLCARD_STATE emStatus;

            public EM_ACCESSCTLCARD_TYPE emType;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string szPsw;

            public int nUseTime;

            public NET_TIME stuValidStartTime;

            public NET_TIME stuValidEndTime;
        }

        public class QueryInfo
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szCardNo;

            public string stuTime;

            public bool bStatus;

            public int nDoor;

            public QueryInfo(string szCardNo, string stuTime, bool bStatus, int nDoor)
            {
                this.szCardNo = szCardNo;
                this.stuTime = stuTime;
                this.bStatus = bStatus;
                this.nDoor = nDoor;
            }
        }

        public class NetworkConfigs
        {
            public string IP { get; set; } 
            public string Mask { get; set; }
            public string GateWay { get; set; }
        }

        public static string SqlBaglantisi = ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
        public CFG_NETWORK_INFO cfg = new CFG_NETWORK_INFO();
        public IntPtr m_LoginID = IntPtr.Zero;
        public NET_DEVICEINFO_Ex m_DeviceInfo;
        public static bool m_IsListen = false;
        public static bool m_bOnline = false;
        public NET_CFG_ACCESS_EVENT_INFO netcfg = new NET_CFG_ACCESS_EVENT_INFO();
        public NET_RECORDSET_ACCESS_CTL_CARD m_stuInfo = new NET_RECORDSET_ACCESS_CTL_CARD();

        private IntPtr m_FindDoorRecordID = IntPtr.Zero;
        private IntPtr m_FindAlarmRecordID = IntPtr.Zero;
        private IntPtr m_FindLogID = IntPtr.Zero;

        internal DahuaIslemler()
        {
            try
            {
                NETClient.Init(null, IntPtr.Zero, null);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(this, ex.Message);
            }
        }

        #region General
        private bool Dahua_GirisYap(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (IntPtr.Zero == m_LoginID)
            {
                ushort port = 0;
                try
                {
                    port = Convert.ToUInt16(CihazPort.Trim());
                }
                catch
                {
                    return false;
                }
                m_DeviceInfo = new NET_DEVICEINFO_Ex();
                m_LoginID = NETClient.LoginWithHighLevelSecurity(CihazIP.Trim(), port, KullaniciAdi.Trim(), Sifre.Trim(), EM_LOGIN_SPAC_CAP_TYPE.TCP, IntPtr.Zero, ref m_DeviceInfo);
                if (IntPtr.Zero == m_LoginID)
                {
                    //MessageBox.Show(this, NETClient.GetLastError());
                    return false;
                }

                m_bOnline = true;
                return true;
            }
            else
            {
                //MessageBox.Show("Can't login: Already logged in");
                return false;
            }
        }

        private bool Dahua_CikisYap()
        {
            if (IntPtr.Zero != m_LoginID)
            {
                bool result = NETClient.Logout(m_LoginID);
                if (!result)
                {
                    //MessageBox.Show(this, NETClient.GetLastError());
                    return false;
                }
                m_LoginID = IntPtr.Zero;
                m_bOnline = false;
                return true;
            }
            else
            {
                //MessageBox.Show("Can't logout: Already logged out");
                return false;
            }
        }

        internal string Dahua_YenidenBaslat(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            IntPtr inPtr = IntPtr.Zero;
            bool ret = NETClient.ControlDevice(m_LoginID, EM_CtrlType.REBOOT, inPtr, 10000);
            if (!ret)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return NETClient.GetLastError();
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        internal string Dahua_Sifirla(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            NET_IN_RESET_SYSTEM stuResetIn = new NET_IN_RESET_SYSTEM();
            stuResetIn.dwSize = (uint)Marshal.SizeOf(typeof(NET_IN_USERINFO_START_FIND));

            NET_OUT_RESET_SYSTEM stuResetOut = new NET_OUT_RESET_SYSTEM();
            stuResetOut.dwSize = (uint)Marshal.SizeOf(typeof(NET_OUT_USERINFO_START_FIND));
            bool nRet = NETClient.ResetSystem(m_LoginID, ref stuResetIn, ref stuResetOut, 5000);
            if (!nRet)
            {
                IntPtr inPtr = IntPtr.Zero;
                inPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NET_RESTORE_TEMPSTRUCT)));
                NET_RESTORE_TEMPSTRUCT temp = new NET_RESTORE_TEMPSTRUCT() { value = NET_RESTORE.ALL };
                Marshal.StructureToPtr(temp, inPtr, true);
                bool ret = NETClient.ControlDevice(m_LoginID, EM_CtrlType.RESTOREDEFAULT, inPtr, 10000);
                if (!ret)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return NETClient.GetLastError();
                }
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }
        #endregion

        #region Door Control
        internal string Dahua_KapiyiAc(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            GetConfig();
            if (netcfg.emState != EM_CFG_ACCESS_STATE.NORMAL)
            {
                netcfg.emState = EM_CFG_ACCESS_STATE.NORMAL;
                SetConfig(netcfg);
            }
            Console.WriteLine();
            NET_CTRL_ACCESS_OPEN openInfo = new NET_CTRL_ACCESS_OPEN();
            openInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_CTRL_ACCESS_OPEN));
            openInfo.nChannelID = Constants.SELECTED_DOOR_INDEX;
            openInfo.szTargetID = IntPtr.Zero;
            IntPtr inPtr = IntPtr.Zero;
            try
            {
                inPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NET_CTRL_ACCESS_OPEN)));
                Marshal.StructureToPtr(openInfo, inPtr, true);
                bool ret = NETClient.ControlDevice(m_LoginID, EM_CtrlType.ACCESS_OPEN, inPtr, 10000);
                if (!ret)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return NETClient.GetLastError();
                }
            }
            finally
            {
                Marshal.FreeHGlobal(inPtr);
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        internal string Dahua_KapiyiKapat(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            GetConfig();
            if (netcfg.emState != EM_CFG_ACCESS_STATE.NORMAL)
            {
                netcfg.emState = EM_CFG_ACCESS_STATE.NORMAL;
                SetConfig(netcfg);
            }
            NET_CTRL_ACCESS_CLOSE closeInfo = new NET_CTRL_ACCESS_CLOSE();
            closeInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_CTRL_ACCESS_CLOSE));
            closeInfo.nChannelID = Constants.SELECTED_DOOR_INDEX;
            IntPtr inPtr = IntPtr.Zero;
            try
            {
                inPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NET_CTRL_ACCESS_CLOSE)));
                Marshal.StructureToPtr(closeInfo, inPtr, true);
                bool ret = NETClient.ControlDevice(m_LoginID, EM_CtrlType.ACCESS_CLOSE, inPtr, 10000);
                if (!ret)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return NETClient.GetLastError();
                }
            }
            finally
            {
                Marshal.FreeHGlobal(inPtr);
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        internal string Dahua_KapiyiAcikTut(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            GetConfig();
            netcfg.emState = EM_CFG_ACCESS_STATE.OPENALWAYS;
            bool bRet = SetConfig(netcfg);
            if (!bRet)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return NETClient.GetLastError();
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        internal string Dahua_KapiyiKapaliTut(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            GetConfig();
            netcfg.emState = EM_CFG_ACCESS_STATE.CLOSEALWAYS;
            bool bRet = SetConfig(netcfg);
            if (!bRet)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return NETClient.GetLastError();
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        public NET_CFG_ACCESS_EVENT_INFO GetConfig()
        {
            try
            {
                object objTemp = new object();
                bool bRet = NETClient.GetNewDevConfig(m_LoginID, Constants.SELECTED_DOOR_INDEX, "AccessControl", ref objTemp, typeof(NET_CFG_ACCESS_EVENT_INFO), 5000);
                if (bRet)
                {
                    netcfg = (NET_CFG_ACCESS_EVENT_INFO)objTemp;
                }
                else
                {
                    //MessageBox.Show(NETClient.GetLastError());
                }
            }
            catch (NETClientExcetion nex)
            {
                //MessageBox.Show(nex.Message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return netcfg;
        }

        public bool SetConfig(NET_CFG_ACCESS_EVENT_INFO cfg)
        {
            bool bRet = false;
            try
            {
                bRet = NETClient.SetNewDevConfig(m_LoginID, Constants.SELECTED_DOOR_INDEX, "AccessControl", (object)cfg, typeof(NET_CFG_ACCESS_EVENT_INFO), 5000);
                if (!bRet)
                {
                    //MessageBox.Show(NETClient.GetLastError());
                }
            }
            catch (NETClientExcetion nex)
            {
                //MessageBox.Show(nex.Message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return bRet;
        }
        #endregion

        #region Device Network Configuration
        internal Object Dahua_AgAyarlariniGetir(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            NetworkConfigs networkConfigs = null;
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            GetConfig_Network();
            for (int i = 0; i < cfg.nInterfaceNum; i++)
            {
                if (string.Equals(System.Text.Encoding.Default.GetString(cfg.szDefInterface), System.Text.Encoding.Default.GetString(cfg.stuInterfaces[i].szName)))
                {
                    if (System.Text.Encoding.Default.GetString(cfg.stuInterfaces[i].szIP)[0] == 0
                        || System.Text.Encoding.Default.GetString(cfg.stuInterfaces[i].szSubnetMask)[0] == 0
                        || System.Text.Encoding.Default.GetString(cfg.stuInterfaces[i].szDefGateway)[0] == 0)
                    {
                        if (!Dahua_CikisYap())
                            return "Logout Error";
                        return "Error while getting network configuration.";
                    }
                    else
                    {
                        networkConfigs = new NetworkConfigs();
                        networkConfigs.IP = System.Text.Encoding.Default.GetString(cfg.stuInterfaces[i].szIP).Replace("\0", "");
                        networkConfigs.Mask = System.Text.Encoding.Default.GetString(cfg.stuInterfaces[i].szSubnetMask).Replace("\0", "");
                        networkConfigs.GateWay = System.Text.Encoding.Default.GetString(cfg.stuInterfaces[i].szDefGateway).Replace("\0", "");
                    }
                }
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return networkConfigs;
        }

        internal string Dahua_AgAyarlariniDegistir(string CihazIP, string IP, string Mask, string GateWay, string CihazPort, string KullaniciAdi, string Sifre)
        {
            string sonuc = Dahua_IPDegistir(CihazIP, IP);
            if (sonuc == "True")
            {
                if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                    return "Login Error";
                if (IP == null || IP == "")
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return "Invalid IP";
                }
                if (Mask == null || Mask == "")
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return "Invalid Mask";
                }
                if (GateWay == null || GateWay == "")
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return "Invalid Gate Way";
                }

                GetConfig_Network();
                for (int i = 0; i < cfg.nInterfaceNum; i++)
                {
                    if (string.Equals("eth0", Encoding.Default.GetString(cfg.stuInterfaces[i].szName).Trim('\0')))
                    {
                        cfg.stuInterfaces[i].szIP = new byte[256];
                        cfg.stuInterfaces[i].szSubnetMask = new byte[256];
                        cfg.stuInterfaces[i].szDefGateway = new byte[256];
                        Encoding.Default.GetBytes(IP.Trim(), 0, IP.Trim().Length, cfg.stuInterfaces[i].szIP, 0);
                        Encoding.Default.GetBytes(Mask.Trim(), 0, Mask.Trim().Length, cfg.stuInterfaces[i].szSubnetMask, 0);
                        Encoding.Default.GetBytes(GateWay.Trim(), 0, GateWay.Trim().Length, cfg.stuInterfaces[i].szDefGateway, 0);
                    }
                }
                bool bRet = SetConfig_Network(cfg);
                if (!bRet)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return NETClient.GetLastError();
                }
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return "True";
            }
            else
            {
                return sonuc;
            }
        }

        private string Dahua_IPDegistir(string CihazIP, string IP)
        {
            if (CihazIP != IP)
            {
                try
                {
                    int sayacCihazGuncelleSonuc = 0;
                    using (SqlConnection con = new SqlConnection(SqlBaglantisi))
                    {
                        con.Open();
                        using (SqlCommand com = new SqlCommand("UPDATE TBLPDKSCihazlar SET IP = @ip WHERE IP = @cihazip", con))
                        {
                            com.Parameters.AddWithValue("@cihazip", CihazIP);
                            com.Parameters.AddWithValue("@ip", IP);
                            sayacCihazGuncelleSonuc = com.ExecuteNonQuery();
                        }
                    }
                    if (sayacCihazGuncelleSonuc > 0)
                    {
                        return "True";
                    }
                    else
                    {
                        return "False";
                    }

                }
                catch (Exception ex)
                {
                    return "Hata Dahua_IPDegistir: \nExp: " + ex;
                }
            }
            else
            {
                return "IP değerleri aynı !";
            }
        }

        private CFG_NETWORK_INFO GetConfig_Network()
        {
            try
            {
                object objTemp = new object();
                bool bRet = NETClient.GetNewDevConfig(m_LoginID, -1, "Network", ref objTemp, typeof(CFG_NETWORK_INFO), 5000);
                if (!bRet)
                {
                    //MessageBox.Show(NETClient.GetLastError());
                    return cfg;
                }
                cfg = (CFG_NETWORK_INFO)objTemp;
            }
            catch (NETClientExcetion nex)
            {
                //MessageBox.Show(nex.Message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return cfg;
        }

        private bool SetConfig_Network(CFG_NETWORK_INFO cfg)
        {
            bool bRet = false;
            try
            {
                bRet = NETClient.SetNewDevConfig(m_LoginID, -1, "Network", (object)cfg, typeof(CFG_NETWORK_INFO), 5000);
            }
            catch (NETClientExcetion nex)
            {
                //MessageBox.Show(nex.Message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return bRet;
        }
        #endregion

        #region Device Time Configuration

        internal string Dahua_ZamaniGetir(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            System.Globalization.CultureInfo cultureinfo = new System.Globalization.CultureInfo("tr-TR");
            NET_TIME stuInfo = new NET_TIME();

            bool ret = NETClient.QueryDeviceTime(m_LoginID, ref stuInfo, 5000);
            if (!ret)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return NETClient.GetLastError();
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return stuInfo.ToDateTime().AddHours(-3.0).ToString(cultureinfo);
        }

        internal string Dahua_ZamaniDegistir(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre, string Zaman)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            System.Globalization.CultureInfo cultureinfo = new System.Globalization.CultureInfo("tr-TR");
            NET_TIME stuSet = new NET_TIME();
            
            stuSet = NET_TIME.FromDateTime(DateTime.Parse(Zaman, cultureinfo).AddHours(3.0));

            bool ret = NETClient.SetupDeviceTime(m_LoginID, stuSet);
            if (!ret)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return NETClient.GetLastError();
            }
            //MessageBox.Show("Set Success");41
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        internal string Dahua_ZamaniSenkronizeEt_PC(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            NET_TIME stuSet = new NET_TIME();
            stuSet = NET_TIME.FromDateTime(DateTime.Now.AddHours(3));

            bool ret = NETClient.SetupDeviceTime(m_LoginID, stuSet);
            if (!ret)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return NETClient.GetLastError();
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        #endregion

        #region Card Management

        internal string Dahua_KartEkle(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre, string CardID, string KullaniciID, string KartSifresi, string KartTipi, string KartKullanimLimiti, string GecerlilikBaslangici, string GecerlilikBitisi)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            // <CARD ID CONVERSION>
            UInt32.TryParse(CardID, out uint parse);
            CardID = Decimal_to_Hex(ReverseBytes(parse).ToString()).PadLeft(8, '0');
            // </CARD ID CONVERSION>

            m_stuInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD));

            NET_CTRL_RECORDSET_INSERT_PARAM stuInfo = new NET_CTRL_RECORDSET_INSERT_PARAM();
            stuInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_CTRL_RECORDSET_INSERT_PARAM));

            stuInfo.stuCtrlRecordSetInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_CTRL_RECORDSET_INSERT_IN));
            stuInfo.stuCtrlRecordSetInfo.emType = EM_NET_RECORD_TYPE.ACCESSCTLCARD;

            stuInfo.stuCtrlRecordSetInfo.nBufLen = (int)Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD));

            stuInfo.stuCtrlRecordSetResult.dwSize = (uint)Marshal.SizeOf(typeof(NET_CTRL_RECORDSET_INSERT_OUT));

            m_stuInfo.stuCreateTime.dwYear = (uint)DateTime.Now.Year;
            m_stuInfo.stuCreateTime.dwMonth = (uint)DateTime.Now.Month;
            m_stuInfo.stuCreateTime.dwDay = (uint)DateTime.Now.Day;
            m_stuInfo.stuCreateTime.dwHour = (uint)DateTime.Now.Hour;
            m_stuInfo.stuCreateTime.dwMinute = (uint)DateTime.Now.Minute;
            m_stuInfo.stuCreateTime.dwSecond = (uint)DateTime.Now.Second;
            m_stuInfo.szCardNo = CardID.Trim();
            m_stuInfo.szUserID = KullaniciID.Trim();
            m_stuInfo.szPsw = KartSifresi.Trim();

            int type = Convert.ToInt32(KartTipi);

            switch (type) // For decreasing the number of card type options
            {
                case 0:
                    type = Constants.CARD_TYPE_GENERAL;
                    break;

                case 1:
                    type = Constants.CARD_TYPE_GUEST;
                    break;

                case 0xFF:
                    type = Constants.CARD_TYPE_MOTHERCARD;
                    break;
            }

            try
            {
                m_stuInfo.emStatus = (EM_ACCESSCTLCARD_STATE)Constants.CARD_STATE_NORMAL;
                m_stuInfo.emType = (EM_ACCESSCTLCARD_TYPE)type;
                m_stuInfo.nUseTime = Convert.ToInt32(KartKullanimLimiti.Trim());
            }
            catch (Exception ex)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return ex.Message;
            }

            // Don't change this region

            #region DontTouch_Door_TimeSection 

            m_stuInfo.bNewDoor = true;

            if (m_stuInfo.nNewDoors == null)
                m_stuInfo.nNewDoors = new int[128];

            m_stuInfo.nNewDoors[0] = 0;
            m_stuInfo.nNewDoorNum = 1;

            if (m_stuInfo.nNewTimeSectionNo == null)
                m_stuInfo.nNewTimeSectionNo = new int[128];

            m_stuInfo.nNewTimeSectionNo[0] = 0;
            m_stuInfo.nNewTimeSectionNum = 1;

            #endregion  

            System.Globalization.CultureInfo cultureinfo = new System.Globalization.CultureInfo("tr-TR");
            DateTime date = DateTime.Parse(GecerlilikBaslangici, cultureinfo);

            m_stuInfo.stuValidStartTime.dwYear = (uint)date.Year;
            m_stuInfo.stuValidStartTime.dwMonth = (uint)date.Month;
            m_stuInfo.stuValidStartTime.dwDay = (uint)date.Day;
            m_stuInfo.stuValidStartTime.dwHour = (uint)date.Hour;
            m_stuInfo.stuValidStartTime.dwMinute = (uint)date.Minute;
            m_stuInfo.stuValidStartTime.dwSecond = (uint)date.Second;

            date = DateTime.Parse(GecerlilikBitisi, cultureinfo);

            m_stuInfo.stuValidEndTime.dwYear = (uint)date.Year;
            m_stuInfo.stuValidEndTime.dwMonth = (uint)date.Month;
            m_stuInfo.stuValidEndTime.dwDay = (uint)date.Day;
            m_stuInfo.stuValidEndTime.dwHour = (uint)date.Hour;
            m_stuInfo.stuValidEndTime.dwMinute = (uint)date.Minute;
            m_stuInfo.stuValidEndTime.dwSecond = (uint)date.Second;
            m_stuInfo.bFirstEnter = Constants.CARD_FIRST_ENTER_STATE;


            IntPtr inPtr = IntPtr.Zero;
            IntPtr ptr = IntPtr.Zero;
            try
            {
                inPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD)));
                Marshal.StructureToPtr(m_stuInfo, inPtr, true);

                stuInfo.stuCtrlRecordSetInfo.pBuf = inPtr;
                stuInfo.stuCtrlRecordSetInfo.nBufLen = (int)Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD));
                ptr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NET_CTRL_RECORDSET_INSERT_PARAM)));
                Marshal.StructureToPtr(stuInfo, ptr, true);
                bool ret = NETClient.ControlDevice(m_LoginID, EM_CtrlType.RECORDSET_INSERT, ptr, 10000);
                if (!ret)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return NETClient.GetLastError();
                }
                stuInfo = (NET_CTRL_RECORDSET_INSERT_PARAM)Marshal.PtrToStructure(ptr, typeof(NET_CTRL_RECORDSET_INSERT_PARAM));
                // MessageBox.Show("Execute Success\n RecNo=" + stuInfo.stuCtrlRecordSetResult.nRecNo.ToString());
            }
            finally
            {
                Marshal.FreeHGlobal(inPtr);
                Marshal.FreeHGlobal(ptr);
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        internal Object Dahua_KartGetir(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre, string KullaniciID)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";

            m_stuInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD));

            int temp;
            if (!int.TryParse(KullaniciID, out temp))
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return "User ID Error";
            }

            NET_CTRL_RECORDSET_PARAM inp = new NET_CTRL_RECORDSET_PARAM();

            //    NET_RECORDSET_ACCESS_CTL_CARD info = new NET_RECORDSET_ACCESS_CTL_CARD();

            IntPtr infoPtr = IntPtr.Zero;
            CardInfo infoList;

            try
            {
                infoPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD)));
                m_stuInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD));
                m_stuInfo.nRecNo = Convert.ToInt32(KullaniciID.Trim());
                m_stuInfo.bEnableExtended = true;

                m_stuInfo.stuFingerPrintInfoEx = new NET_ACCESSCTLCARD_FINGERPRINT_PACKET_EX();
                m_stuInfo.stuFingerPrintInfoEx.nPacketLen = 2000;
                m_stuInfo.stuFingerPrintInfoEx.pPacketData = IntPtr.Zero;
                m_stuInfo.stuFingerPrintInfoEx.pPacketData = Marshal.AllocHGlobal(m_stuInfo.stuFingerPrintInfoEx.nPacketLen);
                Marshal.StructureToPtr(m_stuInfo, infoPtr, true);
                inp.pBuf = infoPtr;
                inp.nBufLen = Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD));
                inp.dwSize = (uint)Marshal.SizeOf(inp);
                inp.emType = EM_NET_RECORD_TYPE.ACCESSCTLCARD;
                object objInp = inp;
                bool ret = NETClient.QueryDevState(m_LoginID, (int)EM_DEVICE_STATE.DEV_RECORDSET_EX, ref objInp, typeof(NET_CTRL_RECORDSET_PARAM), 10000);
                if (!ret)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return NETClient.GetLastError();
                }
                inp = (NET_CTRL_RECORDSET_PARAM)objInp;
                m_stuInfo = (NET_RECORDSET_ACCESS_CTL_CARD)Marshal.PtrToStructure(inp.pBuf, typeof(NET_RECORDSET_ACCESS_CTL_CARD));
                
                try
                {
                    infoList = new CardInfo()
                    {
                        stuCreateTime = m_stuInfo.stuCreateTime,
                        szCardNo = Convert.ToString(Hex_to_Decimal(m_stuInfo.szCardNo)),
                        szUserID = m_stuInfo.szUserID,
                        emType = m_stuInfo.emType,
                        szPsw = m_stuInfo.szPsw,
                        nUseTime = m_stuInfo.nUseTime,
                        stuValidStartTime = m_stuInfo.stuValidStartTime,
                        stuValidEndTime = m_stuInfo.stuValidEndTime
                    };
                }
                catch (Exception exc)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return exc.Message;
                }
            }
            catch (Exception ex)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return ex.Message;
            }
            finally
            {
                Marshal.FreeHGlobal(infoPtr);
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return infoList;
        }

        internal string Dahua_KartSil(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre, string KullaniciID)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            m_stuInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD));
            NET_CTRL_RECORDSET_PARAM stuInfo = new NET_CTRL_RECORDSET_PARAM();
            stuInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_CTRL_RECORDSET_PARAM));
            stuInfo.emType = EM_NET_RECORD_TYPE.ACCESSCTLCARD;
            try
            {
                m_stuInfo.nRecNo = Convert.ToInt32(KullaniciID.Trim());
            }
            catch (Exception ex)
            {
                return ex.Message;
            }


            IntPtr inPtr = IntPtr.Zero;
            IntPtr ptr = IntPtr.Zero;
            try
            {
                inPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(int)));
                Marshal.StructureToPtr(m_stuInfo.nRecNo, inPtr, true);

                stuInfo.pBuf = inPtr;
                stuInfo.nBufLen = (int)Marshal.SizeOf(typeof(int));

                ptr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NET_CTRL_RECORDSET_PARAM)));
                Marshal.StructureToPtr(stuInfo, ptr, true);
                bool ret = NETClient.ControlDevice(m_LoginID, EM_CtrlType.RECORDSET_REMOVE, ptr, 10000);
                if (!ret)
                {
                    return NETClient.GetLastError();
                }
                //MessageBox.Show("Execute Success");
            }
            finally
            {
                Marshal.FreeHGlobal(inPtr);
                Marshal.FreeHGlobal(ptr);
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        internal string Dahua_KartlariTemizle(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";
            m_stuInfo.dwSize = (uint)Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARD));
            NET_CTRL_RECORDSET_PARAM inParam = new NET_CTRL_RECORDSET_PARAM();
            inParam.dwSize = (uint)Marshal.SizeOf(typeof(NET_CTRL_RECORDSET_PARAM));
            inParam.emType = EM_NET_RECORD_TYPE.ACCESSCTLCARD;
            IntPtr inPtr = IntPtr.Zero;
            try
            {
                inPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NET_CTRL_RECORDSET_PARAM)));
                Marshal.StructureToPtr(inParam, inPtr, true);
                bool ret = NETClient.ControlDevice(m_LoginID, EM_CtrlType.RECORDSET_CLEAR, inPtr, 10000);
                if (!ret)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return NETClient.GetLastError();
                }
                //MessageBox.Show("Execute Success");
            }
            catch (Exception ex)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return ex.Message;
            }
            finally
            {
                Marshal.FreeHGlobal(inPtr);
            }
            if (!Dahua_CikisYap())
                return "Logout Error";
            return "True";
        }

        private string Decimal_to_Hex(string decValue)
        {
            return Convert.ToUInt32(decValue).ToString("X");
        }

        private uint Hex_to_Decimal(string hexValue)
        {
            return Convert.ToUInt32(hexValue, 16);
        }

        static uint ReverseBytes(uint val)
        {
            byte[] intAsBytes = BitConverter.GetBytes(val);
            Array.Reverse(intAsBytes);
            return BitConverter.ToUInt32(intAsBytes, 0);
        }

        #endregion

        #region Query Records

        internal Object Dahua_KayitlariGetir(string CihazIP, string CihazPort, string KullaniciAdi, string Sifre, string ZamanBaslangic, string ZamanBitis)
        {
            if (!Dahua_GirisYap(CihazIP, CihazPort, KullaniciAdi, Sifre))
                return "Login Error";

            #region StartQuery
            System.Globalization.CultureInfo cultureinfo = new System.Globalization.CultureInfo("tr-TR");
            if (IntPtr.Zero == m_FindDoorRecordID)
            {
                NET_FIND_RECORD_ACCESSCTLCARDREC_CONDITION_EX condition = new NET_FIND_RECORD_ACCESSCTLCARDREC_CONDITION_EX();
                condition.dwSize = (uint)Marshal.SizeOf(typeof(NET_FIND_RECORD_ACCESSCTLCARDREC_CONDITION_EX));
                condition.bTimeEnable = true;
                condition.stStartTime = NET_TIME.FromDateTime(DateTime.Parse(ZamanBaslangic, cultureinfo));
                condition.stEndTime = NET_TIME.FromDateTime(DateTime.Parse(ZamanBitis, cultureinfo));
                object obj = condition;

                bool ret = NETClient.FindRecord(m_LoginID, EM_NET_RECORD_TYPE.ACCESSCTLCARDREC_EX, obj, typeof(NET_FIND_RECORD_ACCESSCTLCARDREC_CONDITION_EX), ref m_FindDoorRecordID, 10000);
                if (!ret)
                {
                    if (!Dahua_CikisYap())
                        return "Logout Error";
                    return NETClient.GetLastError();
                }
            }
            else
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return "Start Query Error";
            }
            #endregion

            int MaxItemNumber = GetRecordCount();
            int leftItems = MaxItemNumber;
            int hit_offset = -1, miss_offset = -1;
            List<object> ls = new List<object>();
            int retNum = 0;
            for (int i = 0; i < MaxItemNumber; i++)
            {
                NET_RECORDSET_ACCESS_CTL_CARDREC cardrec = new NET_RECORDSET_ACCESS_CTL_CARDREC();
                cardrec.dwSize = (uint)Marshal.SizeOf(typeof(NET_RECORDSET_ACCESS_CTL_CARDREC));
                ls.Add(cardrec);
            }

            List<object> ls2 = new List<object>();

            while(leftItems > 0)
            {
                if(leftItems < 10)
                {
                    NETClient.FindNextRecord(m_FindDoorRecordID, leftItems, ref retNum, ref ls, typeof(NET_RECORDSET_ACCESS_CTL_CARDREC), 10000);
                    leftItems -= leftItems;
                }
                else
                {
                    NETClient.FindNextRecord(m_FindDoorRecordID, 10, ref retNum, ref ls, typeof(NET_RECORDSET_ACCESS_CTL_CARDREC), 10000);
                    leftItems -= 10;
                }
                ls2.AddRange(ls);
            }

            hit_offset = DateFilterHitOffset(ls2, MaxItemNumber, DateTime.Parse(ZamanBitis, cultureinfo));
            miss_offset = DateFilterMissOffset(ls2, MaxItemNumber, DateTime.Parse(ZamanBaslangic, cultureinfo), hit_offset);

            if(miss_offset == hit_offset)
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return "Error: Bu zaman aralığında hiç bir işlem bulunamadı.";
            }
            else
            {
                ls2 = ls2.Skip(hit_offset).Take(miss_offset - hit_offset).ToList();
            }

            List<QueryInfo> ls3 = new List<QueryInfo>();
            foreach (var item in ls2)
            {
                NET_RECORDSET_ACCESS_CTL_CARDREC info = (NET_RECORDSET_ACCESS_CTL_CARDREC)item;
                ls3.Add(new QueryInfo(ReverseBytes(Hex_to_Decimal(info.szCardNo)).ToString().PadLeft(10, '0'), info.stuTime.ToDateTime().ToString(cultureinfo), info.bStatus, info.nDoor));
            }

            #region StopQuery
            if (IntPtr.Zero != m_FindDoorRecordID)
            {
                NETClient.FindRecordClose(m_FindDoorRecordID);
                m_FindDoorRecordID = IntPtr.Zero;
            }
            else
            {
                if (!Dahua_CikisYap())
                    return "Logout Error";
                return "Stop Query Error";
            }
            #endregion

            if (!Dahua_CikisYap())
                return "Logout Error";
            return ls3;
        }

        private int DateFilterMissOffset(List<object> ls, int MaxItemNumber, DateTime ZamanBaslangic, int hit_offset)
        {
            for(int i = MaxItemNumber-1; i>=hit_offset; i--)
            {
                if (DateTime.Compare(((NET_RECORDSET_ACCESS_CTL_CARDREC)ls[i]).stuTime.ToDateTime(), ZamanBaslangic) > 0)
                {
                    return i;
                }
            }
            return -1;
        }

        private int DateFilterHitOffset(List<object> ls, int MaxItemNumber, DateTime ZamanBitis)
        {
            for (int i = 0; i < MaxItemNumber; i++)
            {
                if (DateTime.Compare(((NET_RECORDSET_ACCESS_CTL_CARDREC)ls[i]).stuTime.ToDateTime(), ZamanBitis) <= 0)
                {
                    return i;
                }
            }
            return -1;
        }

        private int GetRecordCount()
        {
            if (IntPtr.Zero == m_FindDoorRecordID)
            {
                return 0;
            }

            int nCount = 0;
            try
            {
                if (NETClient.QueryRecordCount(m_FindDoorRecordID, ref nCount, 3000))
                {
                    return nCount;
                }
                else
                {
                    return 0;
                }
            }
            catch (NETClientExcetion ex)
            {
                return 0;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        #endregion

    }
}