using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Text;
using System.Threading;
using System.Windows;

namespace WpfLibrary.SerialPortBase
{
    public class SerialPortControl : IDisposable
    {
        private SerialPort? _serialPort = null;

        /// <summary>
        /// �����¼��ӳ�ʱ�䣨ms��  
        /// <para>�������ֽڳ���*10/������=��Ҫ�ӳٵ�ʱ��</para>
        /// <para>Ĭ��20ms������20���ֽ� 200/9600Լ����0.0208s</para>
        /// </summary>
        public int ReceiveSleep { get; set; } = 20;

        /// <summary>
        /// ��������
        /// </summary>
        public Encoding Encoding { get; set; } = Encoding.UTF8;

        /// <summary>
        /// ����ί��
        /// </summary>
        public delegate void SerialPortDataReceiveEventArgs(object sender, SerialDataReceivedEventArgs e, byte[] bits);

        /// <summary>
        /// ������������¼�
        /// </summary>
        public event SerialPortDataReceiveEventArgs? DataReceived;

        /// <summary>
        /// �����¼��Ƿ���Ч false��ʾ��Ч
        /// </summary>
        public bool ReceiveEventFlag { get; set; } = false;

        /// <summary>
        /// �Ƿ�򿪴���
        /// </summary>
        public bool IsPortOpen { get { return _serialPort != null && _serialPort.IsOpen; } }

        #region Ĭ�Ϲ��캯��������COM1
        /// <summary>
        /// Ĭ�Ϲ��캯��������COM1
        /// <para>9600 8 N 1</para>
        /// </summary>
        public SerialPortControl()
        {
            _serialPort = new();

            SetSerialPortEvent();
        }

        #endregion

        #region ���캯����comPortName
        /// <summary>
        /// Ĭ�Ϲ��캯��������COM1
        /// <para>9600 8 N 1</para>
        /// </summary>
        /// <param name="comPortName">COM Name</param>
        public SerialPortControl(string comPortName)
        {
            _serialPort = new SerialPort(comPortName)
            {
                BaudRate = 9600,
                Parity = Parity.None,
                DataBits = 8,
                StopBits = StopBits.One,
                //RtsEnable = true,
                //ReadTimeout = 3000,
                Encoding = Encoding
                //���ô���DataReceived�¼����ֽ���Ϊ1
                //ReceivedBytesThreshold = 1
            };
            SetSerialPortEvent();
        }
        #endregion

        #region ���캯��,����comPortName��baudRate
        /// <summary>
        /// ���캯��,����comPortName��baudRate
        /// </summary>
        /// <param name="comPortName">��Ҫ������COM������</param>
        /// <param name="baudRate">COM�Ĳ�����</param>
        public SerialPortControl(string comPortName, int baudRate)
        {
            _serialPort = new SerialPort(comPortName, baudRate)
            {
                Parity = Parity.None,
                DataBits = 8,
                StopBits = StopBits.One,
                Encoding = Encoding
            };
            SetSerialPortEvent();
        }
        #endregion

        #region ���캯��,�����Զ��崮�ڵĳ�ʼ������
        /// <summary>
        /// ���캯��,�����Զ��崮�ڵĳ�ʼ������
        /// <para>comPortName����Ҫ������COM������</para>
        /// <para>baudRate��COM�Ĳ�����</para>
        /// <para>parity����żУ��λ</para>
        /// <para>dataBits������λ</para>
        /// <para>stopBits��ֹͣλ</para>
        /// </summary>
        public SerialPortControl(string comPortName, int baudRate, Parity parity, int dataBits, StopBits stopBits)
        {
            _serialPort = new SerialPort(comPortName, baudRate, parity, dataBits, stopBits)
            {
                Encoding = Encoding
            };
            SetSerialPortEvent();
        }
        #endregion

        #region ˽�����ô����¼�

        /// <summary>
        /// �������ݵĽ���
        /// </summary>
        private void SetSerialPortEvent()
        {
            if (_serialPort == null) return;
            //���յ�һ���ֽ�ʱ��Ҳ�ᴥ��DataReceived�¼�
            _serialPort.DataReceived += SerialPort_DataReceived;
            //�������ݳ���,�����¼�
            _serialPort.ErrorReceived += SerialPort_ErrorReceived;
        }

        #endregion

        #region ���ô��ڲ���

        /// <summary>
        /// ���ô���
        /// <para>comPortName����Ҫ������COM������</para>
        /// <para>baudRate��COM�Ĳ�����</para>
        /// <para>parity����żУ��λ</para>
        /// <para>dataBits������λ</para>
        /// <para>stopBits��ֹͣλ</para>
        /// </summary>
        /// <param name="comPortName"></param>
        /// <param name="baudRate"></param>
        /// <param name="parityBit"></param>
        /// <param name="dataBits"></param>
        /// <param name="stopBits"></param>
        public void SetSerialPort(string comPortName, int baudRate, Parity parityBit, int dataBits, StopBits stopBits)
        {
            if (_serialPort == null) return;
            if (_serialPort.IsOpen)
                _serialPort.Close();
            _serialPort.PortName = comPortName;
            _serialPort.BaudRate = baudRate;
            _serialPort.Parity = parityBit;
            _serialPort.DataBits = dataBits;
            _serialPort.StopBits = stopBits;

            
            SetSerialPortEvent();
        }
        #endregion

        #region ���մ��������¼�
        /// <summary>
        /// ���մ��������¼�
        /// </summary>
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //��ֹ�����¼�ʱֱ���˳�
            if (ReceiveEventFlag) return;
            try
            {
                Thread.Sleep(ReceiveSleep);

                byte[] _data = new byte[_serialPort!.BytesToRead];

                _serialPort.Read(_data, 0, _data.Length);

                if (_data.Length == 0) { return; }

                DataReceived?.Invoke(sender, e, _data);
            }
            catch (Exception ex)
            {

                System.Media.SystemSounds.Beep.Play();
                MessageBox.Show(ex.Message);

            }
        }
        #endregion


        #region �������ݳ����¼�
        /// <summary>
        /// �������ݳ����¼�
        /// </summary>
        private void SerialPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Error");
            // TODO: SerialPort_ErrorReceived
        }
        #endregion

        #region ��������-�ش�ȷ��
        /// <summary>
        /// ��������
        /// 
        /// </summary>
        /// <param name="SendData">��������</param>
        /// <param name="ReceiveData">��������</param>
        /// <param name="Overtime">��ʱʱ�䣬Ĭ��500ms</param>
        /// <returns></returns>
        public bool SendCommand(byte[] SendData, byte[] ReceiveData, int Overtime = 500)
        {
            if (_serialPort!.IsOpen)
            {
                try
                {
                    ReceiveEventFlag = true;        //�رս����¼�

                    _serialPort.DiscardInBuffer();  //��ս��ջ�����                

                    _serialPort.Write(SendData, 0, SendData.Length);

                    int num = 0, outtime = Overtime / 10;

                    //ReceiveEventFlag = false;      //���¼�

                    while (num++ < outtime)
                    {
                        Thread.Sleep(10);

                        if (_serialPort.BytesToRead >= ReceiveData.Length)
                        {
                            byte[] receivedata = new byte[_serialPort.BytesToRead];
                            _serialPort.Read(receivedata, 0, _serialPort.BytesToRead);
                            if (ByteArrayContains(receivedata, ReceiveData))
                            {
                                ReceiveEventFlag = false;      //���¼�
                                return true;
                            }
                        }

                    }

                    ReceiveEventFlag = false;      //���¼�
                    return false;
                }
                catch (Exception ex)
                {
                    ReceiveEventFlag = false;
                    Debug.WriteLine(ex);
                    throw;
                }
            }
            return false;

        }
        #endregion

        #region ��̬����
        #region �Ƚ��ֽڣ������Ƚϣ�
        /// <summary>
        /// Byte�Ƚ�
        /// </summary>
        /// <param name="source">���ж�</param>
        /// <param name="pattern">Ŀ��</param>
        /// <returns></returns>
        public static bool ByteArrayContains(byte[] source, byte[] pattern)
        {
            bool contains = false;
            for (int i = 0; i < source.Length - pattern.Length + 1; i++)
            {
                if (source[i] == pattern[0])
                {
                    bool match = true;
                    for (int j = 1; j < pattern.Length; j++)
                    {
                        if (source[i + j] != pattern[j])
                        {
                            match = false;
                            break;
                        }
                    }
                    if (match)
                    {
                        contains = true;
                        break;
                    }
                }
            }
            return contains;
        }

        #endregion

        #region ʮ�������ַ���ת�ֽ���
        /// <summary>
        /// ��ʮ�������ַ���ת�����ֽ���(����1)
        /// </summary>
        /// <param name="InString"></param>
        /// <returns></returns>
        public static byte[] StringToByte(string InString)
        {
            string[] ByteStrings;
            ByteStrings = InString.Split(" ".ToCharArray());
            byte[] ByteOut;
            ByteOut = new byte[ByteStrings.Length];
            for (int i = 0; i <= ByteStrings.Length - 1; i++)
            {
                //ByteOut[i] = System.Text.Encoding.ASCII.GetBytes(ByteStrings[i]);
                ByteOut[i] = Byte.Parse(ByteStrings[i], System.Globalization.NumberStyles.HexNumber);
                //ByteOut[i] =Convert.ToByte("0x" + ByteStrings[i]);
            }
            return ByteOut;
        }
        #endregion

        #region ʮ�������ַ���ת�ֽ���
        /// <summary>
        /// �ַ���ת16�����ֽ�����(����2)
        /// </summary>
        /// <param name="hexString"></param>
        /// <returns></returns>
        public static byte[] StrToToHexByte(string hexString)
        {
            hexString = hexString.Replace(" ", "");
            if ((hexString.Length % 2) != 0)
                hexString += " ";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            return returnBytes;
        }
        #endregion

        #region �ֽ���תʮ�������ַ���
        /// <summary>
        /// �ֽ�����ת16�����ַ���1
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns>�޿ո�16�����ַ���</returns>
        public static string ByteToHexStr(byte[] bytes)
        {
            string returnStr = "";
            if (bytes != null)
            {
                for (int i = 0; i < bytes.Length; i++)
                {
                    returnStr += bytes[i].ToString("X2");
                }
            }
            return returnStr;
        }
        #endregion

        #region �ֽ���תʮ�������ַ���
        /// <summary>
        /// �ֽ���ת����ʮ�������ַ���2
        /// </summary>
        /// <param name="InBytes"></param>
        /// <returns>���ո�16�����ַ���</returns>
        public static string ByteToString(byte[] InBytes)
        {
            string StringOut = "";
            foreach (byte InByte in InBytes)
            {
                StringOut += String.Format("{0:X2} ", InByte);
            }
            return StringOut;
        }
        #endregion

        #region CRC�㷨16-��ѹ��
        /// <summary>
        /// CRC16�㷨-��ѹ��
        /// </summary>
        /// <param name="hexString"></param>
        /// <returns>string�� ԭ�ֽ�+�����ֽ�</returns>
        public static string CalculateCRC16(string hexString)
        {
            byte[] data = StringToByteArray(hexString);
            ushort crc = 0xFFFF;
            for (int i = 0; i < data.Length; i++)
            {
                crc = (ushort)(crc ^ data[i]);
                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 0x0001) == 1)
                    {
                        crc = (ushort)((crc >> 1) ^ 0xA001);
                    }
                    else
                    {
                        crc = (ushort)(crc >> 1);
                    }
                }
            }
            byte[] crcBytes = BitConverter.GetBytes(crc);
            //Array.Reverse(crcBytes); // ��Ҫ��ת�ֽ���
            string crcString = BitConverter.ToString(crcBytes).Replace("-", "");
            return hexString + crcString;
        }

        /// <summary>
        /// CRC16�㷨
        /// </summary>
        /// <param name="hexString"></param>
        /// <returns>string �����ֽ�</returns>
        public static string GetCRC16(string hexString)
        {
            byte[] data = StringToByteArray(hexString);
            ushort crc = 0xFFFF;
            for (int i = 0; i < data.Length; i++)
            {
                crc = (ushort)(crc ^ data[i]);
                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 0x0001) == 1)
                    {
                        crc = (ushort)((crc >> 1) ^ 0xA001);
                    }
                    else
                    {
                        crc = (ushort)(crc >> 1);
                    }
                }
            }
            byte[] crcBytes = BitConverter.GetBytes(crc);
            //Array.Reverse(crcBytes); // ��Ҫ��ת�ֽ���
            string crcString = BitConverter.ToString(crcBytes).Replace("-", "");
            return crcString;
        }

        /// <summary>
        /// stringתbyte
        /// </summary>
        /// <param name="hexString"></param>
        /// <returns></returns>
        public static byte[] StringToByteArray(string hexString)
        {
            int length = hexString.Length / 2;
            byte[] data = new byte[length];
            for (int i = 0; i < length; i++)
            {
                data[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            }
            return data;
        }
        #endregion

        #region CRC-16/MODBUS�㷨

        /// <summary>
        /// CRC-16/MODBUS�㷨
        /// </summary>
        /// <param name="data"> �������ֽ�</param>
        /// <returns>string �����ֽ�</returns>
        public static string GetCRC16MODBUS(byte[] data)
        {

            ushort crc = 0xFFFF;

            foreach (byte b in data)
            {
                crc ^= b;

                for (int i = 0; i < 8; i++)
                {
                    if ((crc & 0x0001) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else
                    {
                        crc >>= 1;
                    }
                }
            }
            byte[] crcBytes = BitConverter.GetBytes(crc);
            //Array.Reverse(crcBytes); // ��Ҫ��ת�ֽ���
            string crcString = BitConverter.ToString(crcBytes).Replace("-", "");
            return crcString;
        }

        /// <summary>
        /// CRC-16/MODBUS�㷨
        /// </summary>
        /// <param name="data"> �������ֽ�(�ı���)</param>
        /// <returns>string �����ֽ�</returns>
        public static string GetCRC16MODBUS(string hexString)
        {
            byte[] data = StringToByteArray(hexString);
            ushort crc = 0xFFFF;

            foreach (byte b in data)
            {
                crc ^= b;

                for (int i = 0; i < 8; i++)
                {
                    if ((crc & 0x0001) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else
                    {
                        crc >>= 1;
                    }
                }
            }
            byte[] crcBytes = BitConverter.GetBytes(crc);
            //Array.Reverse(crcBytes); // ��Ҫ��ת�ֽ���
            string crcString = BitConverter.ToString(crcBytes).Replace("-", "");
            return crcString;
        }

        #endregion
        #endregion

        #region ע������
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_serialPort != null)
                {
                    if (_serialPort.IsOpen)
                    {
                        _serialPort.Close();
                    }
                    //���յ�һ���ֽ�ʱ��Ҳ�ᴥ��DataReceived�¼�
                    _serialPort.DataReceived -= SerialPort_DataReceived;
                    //�������ݳ���,�����¼�
                    _serialPort.ErrorReceived -= SerialPort_ErrorReceived;
                    _serialPort.Dispose();
                }
            }
        }
        #endregion
    }
}