using System;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Windows.Forms;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Permissions;
using System.Security.Principal;
using System.Management;

namespace COMPortReseter
{
    public partial class Form1 : Form
    {
        public SerialPort serial;
        public List<string> AvailablePorts = new List<string>();     //Available COM Ports List

        public string RegistryKeyAccess = @"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\COM Name Arbiter";
        public string ThisKeyName = "ComDB";

        public byte[] COM_Ports_Bytes = new byte[] { };
        public byte[] ClearData = new byte[] { 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00 };
        public byte[] COM_Ports_Bytes_NEW = new byte[] { };

        public byte[] NewByte = new byte[1] { 0x00 };

        public bool LoadingValues = true;

        public Form1()
        {
            InitializeComponent();

            AdminRelauncher();
            SpawnList();

            LoadingValues = false;
        }

        private void AdminRelauncher()
        {
            if (!IsRunAsAdmin())
            {
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.WorkingDirectory = Environment.CurrentDirectory;
                try
                {
                    proc.FileName = Assembly.GetEntryAssembly().CodeBase;
                }
                catch
                {
                    proc.FileName = Application.ExecutablePath;
                }

                proc.Verb = "runas";

                try
                {
                    Process.Start(proc);
                }
                catch { }
                Environment.Exit(0);
            }
        }

        private bool IsRunAsAdmin()
        {
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(id);

            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        public void GetRegistryKey(string KeyName)
        {
            COM_Ports_Bytes = (byte[]) Registry.GetValue(RegistryKeyAccess, KeyName, ClearData);
            if (COM_Ports_Bytes != null)
            {
                COM_Ports_Bytes_NEW = COM_Ports_Bytes;
                int CurrentByte = 0;
                int CurrentBit = 0;
                LoadingValues = true;
                for (int i = 0; i < 40; i++)
                {
                   if (GetBit(COM_Ports_Bytes[CurrentByte], CurrentBit))
                    {
                        checkedListBox1.SetItemChecked(i, true);
                    }

                    CurrentBit++;
                    if (CurrentBit >= 8)
                    {
                        CurrentByte++;
                        CurrentBit = 0;
                    }
                }
            }
            else
            {
                MessageBox.Show("No Value in COM Name Arbiter");
            }
        }

        public void SetBit(int ByteNubmer, int bitNumber, bool Activated)
        {
            Console.WriteLine("StartByte: " + COM_Ports_Bytes_NEW[ByteNubmer].ToString("X2"));
            BitArray bitArray = new BitArray(new byte[] {COM_Ports_Bytes_NEW[ByteNubmer]});
            bitArray.Set(bitNumber, Activated);
            bitArray.CopyTo(NewByte, 0);

            Console.WriteLine("EndByte: " + COM_Ports_Bytes_NEW[ByteNubmer].ToString("X2"));

            COM_Ports_Bytes_NEW[ByteNubmer] = NewByte[0];
        }

        public bool GetBit(byte b, int bitNumber)
        {
            return (b & (1 << bitNumber)) != 0;
        }

        public void SetRegistryKey(string KeyName)
        {
            try
            {
                Registry.SetValue(RegistryKeyAccess, KeyName, COM_Ports_Bytes_NEW, RegistryValueKind.Binary);
            }
            catch (Exception mess)
            {
                MessageBox.Show("Error setting RegistryKey:" + Environment.NewLine + mess);
            }
        }

        public void RegistryDemandPermission()
        {
            RegistryPermission permission = new RegistryPermission(RegistryPermissionAccess.AllAccess, RegistryKeyAccess);
            permission.Demand();
            permission = null;
        }

        public void SpawnList()
        {
            AvailablePorts.Clear();
            foreach (string s in SerialPort.GetPortNames())
            {
                AvailablePorts.Add(s);
            }

            checkedListBox1.Items.Clear();
            LoadingValues = true;
            for (int i = 0; i < 40; i++)
            {
                string COMName = "COM" + i;
                string BufFinalName = GetPortName(COMName);
                if (BufFinalName != "")
                {
                    checkedListBox1.Items.Add(COMName + " | Connected | " + BufFinalName);
                }
                else
                {
                    checkedListBox1.Items.Add(COMName);
                }
                checkedListBox1.SetItemChecked(i, false);
            }

            //#########################
            RegistryDemandPermission();
            GetRegistryKey(ThisKeyName);
        }

        public string GetPortName(string ThisCOM)
        {
            for (int i = 0; i < AvailablePorts.Count; i++)
            {
                if (AvailablePorts[i] == ThisCOM)
                {
                    return AvailablePorts[i];
                }
            }

            return "";
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            /*if (!LoadingValues)
            {
                int CurrentByte = 0;
                int CurrentBit = 0;
                for (int i = 0; i < 40; i++)
                {
                    SetBit(CurrentByte, CurrentBit, checkedListBox1.GetItemChecked(i));

                    CurrentBit++;
                    if (CurrentBit >= 8)
                    {
                        CurrentByte++;
                        CurrentBit = 0;
                    }
                }

                SetRegistryKey(ThisKeyName);

                SpawnList();    //respawn list
                LoadingValues = false;
            }*/
        }

        private void button1_Click(object sender, EventArgs e)
        {
            COM_Ports_Bytes_NEW = COM_Ports_Bytes;

            RegistryDemandPermission();
            SetRegistryKey(ThisKeyName);
            SpawnList();    //respawn list
            LoadingValues = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!LoadingValues)
            {
                int CurrentByte = 0;
                int CurrentBit = 0;
                for (int i = 0; i < 40; i++)
                {
                    SetBit(CurrentByte, CurrentBit, checkedListBox1.GetItemChecked(i));

                    CurrentBit++;
                    if (CurrentBit >= 8)
                    {
                        CurrentByte++;
                        CurrentBit = 0;
                    }
                }

                RegistryDemandPermission();
                SetRegistryKey(ThisKeyName);

                SpawnList();    //respawn list
                LoadingValues = false;
            }
        }
    }
}
