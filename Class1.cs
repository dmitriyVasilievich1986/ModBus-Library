using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Explicit)]
struct ByteToFloat
{
    [FieldOffset(0)]
    private float Float1;
    [FieldOffset(0)]
    private byte b2;
    [FieldOffset(1)]
    private byte b1;
    [FieldOffset(2)]
    private byte b4;
    [FieldOffset(3)]
    private byte b3;
    public float Out(byte a1, byte a2, byte a3, byte a4)
    {
        b1 = a1;
        b2 = a2;
        b3 = a3;
        b4 = a4;
        return Float1;
    }
}

namespace ModBus_Library
{
    public class ModBus_Libra
    {
        #region Initialization
        public SerialPort Port;
        public int Speed = 115200;
        public string Name = "COM1";

        public delegate void Receive_Handler();
        public event Receive_Handler Receive_Event;
        public delegate void Transmit_Handler();
        public event Transmit_Handler Transmit_Event;

        private ByteToFloat BTF = new ByteToFloat();

        public byte[] Data_Receive;
        public byte[] Data_Transmit;
        public byte[] Data_Interrupt;
        public float[] Result;
        public int Length;
        public int Receive_Length;
        public int Error;
        public bool Exchange = false;
        public string Receive_Array = "";
        public string Transmit_Array = "";

        #endregion

        #region Class Constructor

        public ModBus_Libra(SerialPort port) { Port = port; Port.PortName = Name; Port.BaudRate = Speed; Port.DataBits = 8; Port.DataReceived += new SerialDataReceivedEventHandler(this.Receive); }

        public ModBus_Libra(SerialPort port, string name) { Port = port; Name = name; Port.PortName = Name; Port.BaudRate = Speed; Port.DataBits = 8; Port.DataReceived += new SerialDataReceivedEventHandler(this.Receive); }

        public ModBus_Libra(SerialPort port, Int32 speed) { Port = port; Speed = speed; Port.PortName = Name; Port.BaudRate = Speed; Port.DataBits = 8; Port.DataReceived += new SerialDataReceivedEventHandler(this.Receive); }

        public ModBus_Libra(SerialPort port, string name, Int32 speed) { Port = port; Speed = speed; Name = name; Port.PortName = Name; Port.BaudRate = Speed; Port.DataBits = 8; Port.DataReceived += new SerialDataReceivedEventHandler(this.Receive); }

        #endregion

        public Exception Open()
        {
            if (!Port.IsOpen)
            {
                try
                {
                    Port.PortName = Name;
                    Port.BaudRate = Speed;
                    Port.DataBits = 8;
                    Port.Open();
                    Error = 0;
                }
                catch(Exception err) { return err; }
            }
            return null;
        }

        public void Close() { if (Port.IsOpen) Port.Close(); }

        public void Transmit(byte[] Data_To_Send)
        {
            Exchange = false;
            if (!Port.IsOpen) return;
            Data_Transmit = new byte[Data_To_Send.Length + 2];

            if (Data_Interrupt != null) Data_Transmit = Data_Interrupt;
            else
            {
                Data_Transmit[Data_To_Send.Length] = (byte)ModRTU_CRC(Data_To_Send, Data_To_Send.Length);
                Data_Transmit[Data_To_Send.Length + 1] = (byte)(ModRTU_CRC(Data_To_Send, Data_To_Send.Length) >> 8);
                for (int a = 0; a < Data_To_Send.Length; a++) Data_Transmit[a] = Data_To_Send[a];
            }

            if (Data_Transmit[1] == 0x02 && Data_Transmit.Length > 6) 
            {
                Receive_Length = (Data_Transmit[5] - 1) / 8 + 6;
                Length = 0;
            }
            else if (Data_Transmit[1] == 0x03 && Data_Transmit.Length > 6)
            {
                Receive_Length = Data_Transmit[5] * 2 + 5;
                Length = Data_Transmit[5] / 2;
            }
            else if (Data_Transmit[1] == 0x04 && Data_Transmit.Length > 6)
            {
                Receive_Length = Data_Transmit[5] * 2 + 5;
                Length = Data_Transmit[5] / 2;
            }
            else
            {
                Receive_Length = 8;
                Length = 0;
            }

            try
            {
                Port.DiscardInBuffer();
                Port.Write(Data_Transmit, 0, Data_Transmit.Length);
            }
            catch (Exception) { return; }

            Transmit_Array = "Передача: ";
            foreach (byte a in Data_Transmit) Transmit_Array += a.ToString("X2") + " ";

            Transmit_Event?.Invoke();
        }

        public void Interrupt(byte[] Data_To_Send)
        {
            if (!Port.IsOpen || Data_Interrupt != null) return;

            Data_Interrupt = new byte[Data_To_Send.Length + 2];

            for (int a = 0; a < Data_To_Send.Length; a++) Data_Interrupt[a] = Data_To_Send[a];
            Data_Interrupt[Data_To_Send.Length] = (byte)ModRTU_CRC(Data_To_Send, Data_To_Send.Length);
            Data_Interrupt[Data_To_Send.Length + 1] = (byte)(ModRTU_CRC(Data_To_Send, Data_To_Send.Length) >> 8);            
        }

        public void Receive(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            Exchange = true;

            if (Port.BytesToRead < Receive_Length || Port.BytesToRead < 3)   { return; }

            Data_Receive = new byte[Receive_Length > 3 ? Receive_Length : 8];
            try
            {
                Port.Read(Data_Receive, 0, Receive_Length);
                Port.DiscardInBuffer();
            }
            catch (Exception) { return; }
            
            Result = new float[Length + 1];

            if (ModRTU_CRC(Data_Receive, Data_Receive.Length - 2) != (short)((short)Data_Receive[Data_Receive.Length - 2] | ((short)Data_Receive[Data_Receive.Length - 1] << 8))
                || Data_Transmit[0] != Data_Receive[0] || Data_Transmit[1] != Data_Receive[1])
            { Error++; return; }

            switch (Data_Receive[1])
            {
                case 0x02:
                    if (Data_Receive[2] != ((Data_Transmit[5] - 1) / 8 + 1)) { Error++; return; }
                    if (Data_Interrupt != null && Data_Interrupt[1] == 0x02) Data_Interrupt = null;
                    break;
                case 0x03:
                    if (Data_Receive[2] != (Data_Transmit[5] * 2)) { Error++; return; }
                    if (Data_Interrupt != null && Data_Interrupt[1] == 0x03) Data_Interrupt = null;
                    break;
                case 0x04:
                    if (Data_Receive[2] != (Data_Transmit[5] * 2)) { Error++; return; }
                    if (Data_Interrupt != null && Data_Interrupt[1] == 0x04) Data_Interrupt = null;
                    break;
                case 0x06:
                    for (int count = 0; count < Data_Transmit.Length; count++)
                    {
                        if (Data_Receive[count] != Data_Transmit[count]) { Error++; return; }
                    }
                    if (Data_Interrupt != null && Data_Interrupt[1] == 0x06) Data_Interrupt = null;
                    break;
                case 0x10:
                    for (int count = 0; count < 6; count++)
                    {
                        if (Data_Receive[count] != Data_Transmit[count]) { Error++; return; }
                    }
                    if (Data_Interrupt != null && Data_Interrupt[1] == 0x10) Data_Interrupt = null;
                    break;
            }

            for (int x = 0; x < Length; x++)
            {
                Result[x] = BTF.Out(Data_Receive[3 + x * 4], Data_Receive[4 + x * 4], Data_Receive[5 + x * 4], Data_Receive[6 + x * 4]);
            }            

            Receive_Array = "Прием: ";
            foreach (byte a in Data_Receive) Receive_Array += a.ToString("X2") + " ";

            Receive_Event?.Invoke();
        }

        private short ModRTU_CRC(byte[] data, int count)
        {
            ulong crc = 0xFFFF;
            for (int pos = 0; pos < count; pos++)
            {
                crc ^= (ulong)data[pos];
                for (int i = 0; i < 8; i++)
                {
                    if ((crc & 1) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else crc >>= 1;
                }
            }
            return (short)crc;
        }
    }
}
