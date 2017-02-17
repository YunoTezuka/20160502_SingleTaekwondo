using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace God
{
	public static class Yuan {
		static List<string> PathList = new List<string>();
		static List<string> WriteFileBuf = new List<string>();
		static List<DateTime> timestamp = new List<DateTime>();
		public static string[] GetCOMPorts() {
			//System.Windows.Forms.ComboBox combobox = new ComboBox();
			string[] ComPortList = System.IO.Ports.SerialPort.GetPortNames();	//抓目前有連接裝置的COMPort
			string SwapBuf;
			int tmp;
			for(int i = 0; i < ComPortList.Length; i++) {						//放進combobox並排序
				if(!char.IsDigit(ComPortList[i][ComPortList[i].Length - 1])) {	//如果最後一位不是數字就把最後一個字元刪掉，framework 4.0以上好像不會有問題
					ComPortList[i] = ComPortList[i].Substring(0, ComPortList[i].Length - 1);
				}
				//if (i > 0){  //排序，由小排到大
				tmp = i;
				for(int j = 0; j < i; j++) {
					if(Convert.ToInt32(ComPortList[i].Substring(3)) < Convert.ToInt32(ComPortList[j].Substring(3))) {
						SwapBuf = ComPortList[j];
						ComPortList[j] = ComPortList[i];
						ComPortList[i] = SwapBuf;
					}
				}
				//combobox.Items.Insert(tmp, ComPortList[i]);
				//}
				//else combobox.Items.Add(ComPortList[i]);
			}
			return ComPortList;
		}
		public static bool PackageCheck(byte[] RevBuf, byte[] PackageHeader, byte[] PackageTail, int PackageLength) {
			bool isPackage = true;
			for(int i = 0; i < PackageHeader.Length; i++) {
				if(RevBuf[i] != (byte) PackageHeader[i]) {
					isPackage = false;
					break;
				}
			}
			if(isPackage) {
				for(int i = PackageLength - PackageTail.Length; i < PackageLength; i++) {
					if(RevBuf[i] != (byte) PackageTail[i - PackageLength + PackageTail.Length]) {
						isPackage = false;
						break;
					}
				}
			}
			return isPackage;
		}
		public static void SerialportReceiveQueue(System.IO.Ports.SerialPort Serialport, ref System.Collections.Queue RecQueue) {
			int BytesCont = Serialport.BytesToRead; //將全部收到的資料都塞到Queue，因為一直去讀取IO會很花時間
			byte[] RecBuf = new byte[BytesCont];
			try {
				Serialport.Read(RecBuf, 0, BytesCont);
				for(int i = 0; i < BytesCont; i++)
					RecQueue.Enqueue(RecBuf[i]);
			} catch {
				throw;
			}

		}
		public static void CheckAndDisconnection(System.IO.Ports.SerialPort Spt) {
			if(Spt.IsOpen) Spt.Close();
		}
		public static void CheckAndConnection(System.IO.Ports.SerialPort Spt, string PortName, ref string SavePort) {
			try {
				if(!Spt.IsOpen) {
					Spt.PortName = PortName;
					Spt.Open();
					SavePort = PortName;
				}
			} catch {
				throw;
			}

		}
		public static void CheckAndConnection(System.IO.Ports.SerialPort Spt, string PortName) {
			if(!Spt.IsOpen) {
				Spt.PortName = PortName;
				Spt.Open();
			}
		}
		public static void CheckAndClear(System.IO.Ports.SerialPort Spt) {
			if(Spt.IsOpen) {
				string ReadBuf = Spt.ReadExisting();
			}
		}
		/// <summary>
		/// weight of Var = 1/(ESALPHA + 1)
		/// </summary>
		/// <param name="VarES"></param>
		/// <param name="Var"></param>
		/// <param name="ESALPHA"></param>
		public static void ExponentialSmooth(ref float VarES, float Var, uint ESALPHA) {
			VarES = Var / (ESALPHA + 1) + VarES * ESALPHA / (ESALPHA + 1);
		}

		public static void ExponentialSmooth(ref double VarES, float Var, uint ESALPHA) {
			VarES = Var / (ESALPHA + 1) + VarES * ESALPHA / (ESALPHA + 1);
		}
		/// <summary>
		/// if find, return first index
		/// else     return -1   
		/// </summary>
		/// <param name="Source"></param>
		/// <param name="Target"></param>
		/// <returns></returns>
		public static int ByteArrayIndexofByteArray(byte[] Source, byte[] Target) {
			int[] Index = new int[Target.Length];
			int i, j;
			for(i = 0; i <= Source.Length - Target.Length; i++) {
				for(j = 0; j < Target.Length; j++) {
					if(Source[i + j] != Target[j]) break;
				}
				if(j == Target.Length) break;
			}
			return (i > Source.Length - Target.Length) ? -1 : i;
		}
		public static void WriteFile(string path, string Text) {
			int index = PathList.IndexOf(path);
			DateTime now = DateTime.Now;
			if(index < 0) {
				PathList.Add(path);
				timestamp.Add(now);
				WriteFileBuf.Add(Text);
			} else {
				timestamp[index] = now;
				WriteFileBuf[index] += Text;
				if(((TimeSpan) (now - timestamp[index])).TotalMilliseconds > 33) {
					timestamp[index] = now;
					WriteFileBuf[index] += Text;
					using(System.IO.StreamWriter file = new System.IO.StreamWriter(path)) {
						file.Write(WriteFileBuf[index]);
					}
					WriteFileBuf[index] = "";
				}
			}
		}
		public static void PackageShiftAndGetNew(ref byte[] raw, int index, ref System.Collections.Queue RecQueue) {
			if(index == 0) {
				for(int i = 0; i < raw.Length; i++)
					raw[i] = Convert.ToByte(RecQueue.Dequeue());
			} else if(index > 0) {
				Array.Copy(raw, index, raw, 0, raw.Length - index);
				for(int i = index; i > 0; i--)
					raw[raw.Length - i] = Convert.ToByte(RecQueue.Dequeue());
			} else {
				for(int i = 0; i < raw.Length; i++)
					raw[i] = Convert.ToByte(RecQueue.Dequeue());
			}
		}
		//~~~~~~~~~~~~~~~New~~~~~~~~~~~~~
		public static int GetIndexofObjectArray(object[] Array, object Target) {
			for(int i = 0; i < Array.GetLength(0); i++) {
				if(Array[i].Equals(Target)) return i;
			}
			return -1;
		}
		public static int ArrayGetMinIndex<T>(T[] Array) {
			return System.Array.IndexOf(Array, Array.Min());
		}
		public static int ArrayGetMaxIndex<T>(T[] Array) {
			return System.Array.IndexOf(Array, Array.Max());
		}
		public static void LblAutoFontSize_SingleLine(Label Lbl, int MaxWidth, int MaxFontSize) {
			Lbl.AutoSize = true;
			Lbl.Font = new Font(Lbl.Font.FontFamily, MaxFontSize);
			while(Lbl.Width > MaxWidth) Lbl.Font = new Font(Lbl.Font.FontFamily, Math.Max(Lbl.Font.Size - 1, 0.000000001f));
			Lbl.AutoSize = false;
		}
		public static void LblAutoFontSize_SingleLine(Label Lbl) {
			Lbl.SuspendLayout();
			int MaxWidth = Lbl.Width, MaxHeight = Lbl.Height;
			Lbl.AutoSize = true;
			Lbl.Margin = new System.Windows.Forms.Padding(0);
			Lbl.Font = new Font(Lbl.Font.FontFamily, 300, Lbl.Font.Style);
			Lbl.Font = new Font(Lbl.Font.FontFamily, 300 * 1.2f / ((Lbl.Width + 1f) / (MaxWidth + 1f)), Lbl.Font.Style);
			while(Lbl.Width > MaxWidth || Lbl.Height > MaxHeight) Lbl.Font = new Font(Lbl.Font.FontFamily, Math.Max(Lbl.Font.Size - 1, 0.000000001f), Lbl.Font.Style);
			Lbl.AutoSize = false;
			Lbl.Size = new Size(MaxWidth, MaxHeight);
			Lbl.ResumeLayout();
		}
		public static void LblAutoFontSize(Label Lbl, int MaxFontSize) {
			float FontSize = MaxFontSize;
			float Xcnt = Convert.ToSingle((Lbl.Width - 2 * FontSize) / 1.4f / FontSize),
				  YCnt = Convert.ToSingle(Lbl.Height / 1.4f / FontSize);
			float TotalCnt = Xcnt * YCnt;
			if(Lbl.Text.Length > 0) {
				do {
					FontSize = (TotalCnt != 0) ? FontSize / Convert.ToSingle(Math.Sqrt(Lbl.Text.Length / TotalCnt)) : FontSize - 1;
					Xcnt = (Lbl.Width - 2 * FontSize) / 1.4f / FontSize;
					YCnt = Lbl.Height / 1.4f / FontSize;
					TotalCnt = Xcnt * YCnt;
				} while(TotalCnt < Lbl.Text.Length);
				Lbl.Font = new Font(Lbl.Font.FontFamily, FontSize);
			}
		}
			//~~~~~~~~~~~~~End~~~~~~~~~~~~~~~~~~
	}
}
