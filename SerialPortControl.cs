using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Ports;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;

namespace WpfLibrary.SerialPortBase
{
    public class SerialPortControl : IDisposable
    {
        private SerialPort? _serialPort = null;

        #region 公开字段
        /// <summary>
        /// 接收事件延迟时间（ms）  
        /// <para>欲接收字节长度*10/波特率=需要延迟的时间</para>
        /// <para>默认20ms，例：20个字节 200/9600约等于0.0208s</para>
        /// </summary>
        public int ReceiveSleep { get; set; } = 20;

        /// <summary>
        /// 编码类型
        /// </summary>
        public Encoding Encoding { get; set; } = Encoding.UTF8;

        /// <summary>
        /// 定义委托
        /// </summary>
        public delegate void SerialPortDataReceiveEventArgs(object sender, SerialDataReceivedEventArgs e, byte[] bits);

        /// <summary>
        /// 定义接收数据事件
        /// </summary>
        public event SerialPortDataReceiveEventArgs? DataReceived;

        /// <summary>
        /// 接收事件是否有效 false表示有效
        /// </summary>
        public bool ReceiveEventFlag { get; set; } = false;

        /// <summary>
        /// 是否打开串口
        /// </summary>
        public bool IsPortOpen { get { return _serialPort != null && _serialPort.IsOpen; } }
        #endregion

        #region 构造方法
        #region 默认构造函数，操作COM1
        /// <summary>
        /// 默认构造函数，操作COM1
        /// <para>9600 8 N 1</para>
        /// </summary>
        public SerialPortControl()
        {
            _serialPort = new();

            SetSerialPortEvent();
        }

        #endregion

        #region 构造函数，comPortName
        /// <summary>
        /// 默认构造函数，操作COM1
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
                //设置触发DataReceived事件的字节数为1
                //ReceivedBytesThreshold = 1
            };
            SetSerialPortEvent();
        }
        #endregion

        #region 构造函数,操作comPortName，baudRate
        /// <summary>
        /// 构造函数,操作comPortName，baudRate
        /// </summary>
        /// <param name="comPortName">需要操作的COM口名称</param>
        /// <param name="baudRate">COM的波特率</param>
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

        #region 构造函数,可以自定义串口的初始化参数
        /// <summary>
        /// 构造函数,可以自定义串口的初始化参数
        /// <para>comPortName：需要操作的COM口名称</para>
        /// <para>baudRate：COM的波特率</para>
        /// <para>parity：奇偶校验位</para>
        /// <para>dataBits：数据位</para>
        /// <para>stopBits：停止位</para>
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
        #endregion

        #region 私有方法

        /// <summary>
        /// 设置数据的接收
        /// </summary>
        private void SetSerialPortEvent()
        {
            if (_serialPort == null) return;
            //接收到一个字节时，也会触发DataReceived事件
            _serialPort.DataReceived += SerialPort_DataReceived;
            //接收数据出错,触发事件
            _serialPort.ErrorReceived += SerialPort_ErrorReceived;
        }

        #endregion

        #region 公开方法
        #region 打开串口资源
        /// <summary>
        /// 打开串口资源
        /// <returns>返回bool类型</returns>
        /// </summary>
        public bool OpenPort()
        {
            bool ok = false;
            if (_serialPort == null) return ok;
            //如果串口是打开的，先关闭
            if (_serialPort.IsOpen)
                _serialPort.Close();
            //打开串口
            _serialPort.Open();
            ok = _serialPort.IsOpen;
            return ok;
        }
        #endregion

        #region 关闭串口资源
        /// <summary>
        /// 关闭串口资源,操作完成后,一定要关闭串口
        /// </summary>
        public void ClosePort()
        {
            //如果串口处于打开状态,则关闭
            if (_serialPort != null && _serialPort.IsOpen)
                _serialPort.Close();
        }
        #endregion
        /// <summary>
        /// 设置串口
        /// <para>comPortName：需要操作的COM口名称</para>
        /// <para>baudRate：COM的波特率</para>
        /// <para>parity：奇偶校验位</para>
        /// <para>dataBits：数据位</para>
        /// <para>stopBits：停止位</para>
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
            //SetSerialPortEvent();
        }
        #endregion

        #region 接受事件
        #region 接收串口数据事件
        /// <summary>
        /// 接收串口数据事件
        /// </summary>
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //禁止接收事件时直接退出
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

        #region 接收数据出错事件
        /// <summary>
        /// 接收数据出错事件
        /// </summary>
        private void SerialPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Error");
            // TODO: SerialPort_ErrorReceived
        }
        #endregion
        #endregion

        #region 发送方法
        #region 发送数据string类型
        /// <summary>
        /// 发送文本型数据 默认：ASCII
        /// </summary>
        /// <param name="data">待发送的数据</param>
        /// <param name="ishexstring">ture时，以Hex模式发送</param>
        public void SendData(string data, bool ishexstring = false)
        {
            //发送数据
            //禁止接收事件时直接退出
            if (ReceiveEventFlag)
            {
                return;
            }
            try
            {
                if (_serialPort!.IsOpen)
                {
                    if (!ishexstring)
                    {
                        _serialPort.Write(data);
                    }
                    else
                    {
                        // 去掉空格、换行符和 "0x" 前缀
                        string send_data = Regex.Replace(data, @"\s|0x", "");
                        // 将字符串转换成 byte 数组
                        byte[] bytes = new byte[(send_data.Length + 1) / 2];
                        for (int i = 0; i < send_data.Length; i += 2)
                        {
                            bytes[i / 2] = Convert.ToByte(send_data.Substring(i, Math.Min(2, send_data.Length - i)), 16);
                        }

                        // 发送数据
                        _serialPort.Write(bytes, 0, bytes.Length);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Media.SystemSounds.Beep.Play();
                Debug.WriteLine(ex.Message);
                throw;
            }

        }
        #endregion

        #region 发送数据byte类型
        /// <summary>
        /// 发送字节型数据
        /// </summary>
        /// <param name="data">待发送的数据</param>
        /// <param name="offset">从第几位开始发送</param>
        /// <param name="count">发送长度</param>
        public void SendData(byte[] data, int offset, int count)
        {
            //禁止接收事件时直接退出
            if (ReceiveEventFlag)
            {
                return;
            }
            try
            {
                if (_serialPort!.IsOpen)
                {
                    //_serialPort.DiscardInBuffer();//清空接收缓冲区
                    _serialPort.Write(data, offset, count);
                }
            }
            catch (Exception ex)
            {
                System.Media.SystemSounds.Beep.Play();
                Debug.WriteLine(ex.Message);
                throw;
            }
        }
        #endregion

        #region 发送命令-回传确认-byte
        /// <summary>
        /// 发送命令
        /// byte全对比版本
        /// </summary>
        /// <param name="SendData">发送数据</param>
        /// <param name="ReceiveData">接收数据</param>
        /// <param name="Overtime">超时时间，默认500ms</param>
        /// <returns></returns>
        public bool SendCommand(byte[] SendData, byte[] ReceiveData, int Overtime = 500)
        {
            if (_serialPort!.IsOpen)
            {
                try
                {
                    ReceiveEventFlag = true;        //关闭接收事件

                    _serialPort.DiscardInBuffer();  //清空接收缓冲区                

                    _serialPort.Write(SendData, 0, SendData.Length);

                    int num = 0, outtime = Overtime / 10;

                    //ReceiveEventFlag = false;      //打开事件

                    while (num++ < outtime)
                    {
                        Thread.Sleep(10);

                        if (_serialPort.BytesToRead >= ReceiveData.Length)
                        {
                            byte[] receivedata = new byte[_serialPort.BytesToRead];
                            _serialPort.Read(receivedata, 0, _serialPort.BytesToRead);
                            if (ByteArrayContains(receivedata, ReceiveData))
                            {
                                ReceiveEventFlag = false;      //打开事件
                                return true;
                            }
                        }

                    }

                    ReceiveEventFlag = false;      //打开事件
                    return false;
                }
                catch (Exception ex)
                {
                    ReceiveEventFlag = false;
                    Debug.WriteLine(ex.Message);
                    throw;
                }
            }
            return false;

        }
        #endregion
        #region 发送命令-回传确认-string
        /// <summary>
        /// 发送命令
        /// string对比版本 默认全相等对比
        /// </summary>
        /// <param name="SendData">发送数据</param>
        /// <param name="ReceiveData">接收数据</param>
        /// <param name="Overtime">超时时间，默认500ms</param>
        /// <param name="AllContrast">是否全对比，默认全对比</param>
        /// <returns></returns>
        public bool SendCommand(string SendData, string ReceiveData, int Overtime = 500, bool AllContrast = true)
        {
            if (_serialPort!.IsOpen)
            {
                try
                {
                    ReceiveEventFlag = true;        //关闭接收事件

                    _serialPort.DiscardInBuffer();  //清空接收缓冲区                

                    _serialPort.Write(SendData);

                    int num = 0, outtime = Overtime / 10;

                    int len = Encoding.UTF8.GetBytes(ReceiveData).Length;

                    while (num++ < outtime)
                    {
                        Thread.Sleep(10);

                        if (_serialPort.BytesToRead >= len)
                        {
                            byte[] receivedata = new byte[_serialPort.BytesToRead];
                            _serialPort.Read(receivedata, 0, _serialPort.BytesToRead);
                            // 接收文本
                            string recdata = Encoding.Default.GetString(receivedata);

                            // 判断相等
                            if (AllContrast)
                            {
                                // 不相等 继续循环
                                if (!string.Equals(recdata, ReceiveData, StringComparison.Ordinal))
                                    continue;
                            }
                            else
                            {
                                if (!recdata.Contains(ReceiveData))
                                    continue;
                            }
                            ReceiveEventFlag = false;      //打开事件
                            return true;
                        }

                    }

                    ReceiveEventFlag = false;      //打开事件
                    return false;
                }
                catch (Exception ex)
                {
                    ReceiveEventFlag = false;
                    Debug.WriteLine(ex.Message);
                    throw;
                }
            }
            return false;

        }
        #endregion
        #endregion

        #region 静态方法

        /// <summary>
        /// 获取可用串口名称
        /// </summary>
        /// <returns></returns>
        public static List<string> GetPortNames() => new(SerialPort.GetPortNames());

        #region 比较字节（暴力比较）
        /// <summary>
        /// Byte比较
        /// </summary>
        /// <param name="source">待判断</param>
        /// <param name="pattern">目标</param>
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

        #region 十六进制字符串转字节型
        /// <summary>
        /// 把十六进制字符串转换成字节型(方法1)
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
                ByteOut[i] = byte.Parse(ByteStrings[i], System.Globalization.NumberStyles.HexNumber);
                //ByteOut[i] =Convert.ToByte("0x" + ByteStrings[i]);
            }
            return ByteOut;
        }
        #endregion

        #region 十六进制字符串转字节型
        /// <summary>
        /// 字符串转16进制字节数组(方法2)
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

        #region 十六进制字符串转字节型
        /// <summary>
        /// string转byte
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

        #region 字节型转十六进制字符串
        /// <summary>
        /// 字节数组转16进制字符串1
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns>无空格16进制字符串</returns>
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

        #region 字节型转十六进制字符串
        /// <summary>
        /// 字节型转换成十六进制字符串2
        /// </summary>
        /// <param name="InBytes"></param>
        /// <returns>带空格16进制字符串</returns>
        public static string ByteToString(byte[] InBytes)
        {
            string StringOut = "";
            foreach (byte InByte in InBytes)
            {
                StringOut += string.Format("{0:X2} ", InByte);
            }
            return StringOut;
        }
        #endregion

        #region CRC算法16-电压器
        /// <summary>
        /// CRC16算法-电压器
        /// </summary>
        /// <param name="hexString"></param>
        /// <returns>string型 原字节+计算字节</returns>
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
            //Array.Reverse(crcBytes); // 需要翻转字节序
            string crcString = BitConverter.ToString(crcBytes).Replace("-", "");
            return hexString + crcString;
        }

        /// <summary>
        /// CRC16算法
        /// </summary>
        /// <param name="hexString"></param>
        /// <returns>string 计算字节</returns>
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
            //Array.Reverse(crcBytes); // 需要翻转字节序
            string crcString = BitConverter.ToString(crcBytes).Replace("-", "");
            return crcString;
        }


        #endregion

        #region CRC-16/MODBUS算法

        /// <summary>
        /// CRC-16/MODBUS算法
        /// </summary>
        /// <param name="data"> 欲计算字节</param>
        /// <returns>string 计算字节</returns>
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
            //Array.Reverse(crcBytes); // 需要翻转字节序
            string crcString = BitConverter.ToString(crcBytes).Replace("-", "");
            return crcString;
        }

        /// <summary>
        /// CRC-16/MODBUS算法
        /// </summary>
        /// <param name="data"> 欲计算字节(文本型)</param>
        /// <returns>string 计算字节</returns>
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
            //Array.Reverse(crcBytes); // 需要翻转字节序
            string crcString = BitConverter.ToString(crcBytes).Replace("-", "");
            return crcString;
        }

        #endregion

        #region 累加和校验
        /// <summary>
        /// 累加和校验
        /// </summary>
        /// <param name="data"></param>
        /// <returns>byte</returns>
        public static byte SumToByte(byte[] data)
        {
            byte sum = 0;
            foreach (byte b in data)
            {
                sum += b;
            }
            return sum;
        }

        /// <summary>
        /// 累加和校验
        /// </summary>
        /// <param name="hexString">十六进制文本</param>
        /// <returns>byte</returns>
        public static byte SumToByte(string hexString)
        {
            byte[] data = StringToByteArray(hexString);
            byte sum = 0;
            foreach (byte b in data)
            {
                sum += b;
            }
            return sum;
        }

        /// <summary>
        /// 累加和校验
        /// </summary>
        /// <param name="data"></param>
        /// <returns>string</returns>
        public static string SumToString(byte[] data)
        {
            byte sum = 0;
            foreach (byte b in data)
            {
                sum += b;
            }
            return sum.ToString("X2");
        }

        /// <summary>
        /// 累加和校验
        /// </summary>
        /// <param name="hexString">十六进制文本</param>
        /// <returns>string</returns>
        public static string SumToString(string hexString)
        {
            byte[] data = StringToByteArray(hexString);
            byte sum = 0;
            foreach (byte b in data)
            {
                sum += b;
            }
            return sum.ToString("X2");
        }
        #endregion

        #region 只允许数字和ABCDEF
        /// <summary>
        /// 正则表达式判断只允许数字和ABCDEF
        /// </summary>
        /// <param name="text"></param>
        /// <returns>true为不符合 false为符合</returns>
        public static bool IsTextAllowed(string text)
        {
            Regex regex = new("[0-9a-fA-F]+$"); // 只允许数字和ABCDEF
            return regex.IsMatch(text);
        }
        #endregion
        #endregion

        #region 注销方法
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
                    //接收到一个字节时，也会触发DataReceived事件
                    _serialPort.DataReceived -= SerialPort_DataReceived;
                    //接收数据出错,触发事件
                    _serialPort.ErrorReceived -= SerialPort_ErrorReceived;
                    _serialPort.Dispose();
                }
            }
        }
        #endregion
    }
}
