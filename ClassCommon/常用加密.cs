using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ClassCommon
{
   public class 常用加密
    {
        private string MD5(string val)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] palindata = Encoding.Default.GetBytes(val);//将要加密的字符串转换为字节数组
            byte[] encryptdata = md5.ComputeHash(palindata);//将字符串加密后也转换为字符数组
            return Convert.ToBase64String(encryptdata);//将加密后的字节数组转换为加密字符串
        }

        //RSA加密算法
        private string Encryption(string express)
        {
            CspParameters param = new CspParameters();
            param.KeyContainerName = "oa_erp_dowork";//密匙容器的名称，保持加密解密一致才能解密成功
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(param))
            {
                byte[] plaindata = Encoding.Default.GetBytes(express);//将要加密的字符串转换为字节数组
                byte[] encryptdata = rsa.Encrypt(plaindata, false);//将加密后的字节数据转换为新的加密字节数组
                return Convert.ToBase64String(encryptdata);//将加密后的字节数组转换为字符串
            }
        }
        //RSA解密
        private string Decrypt(string ciphertext)
        {
            CspParameters param = new CspParameters();
            param.KeyContainerName = "oa_erp_dowork";
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(param))
            {
                byte[] encryptdata = Convert.FromBase64String(ciphertext);
                byte[] decryptdata = rsa.Decrypt(encryptdata, false);
                return Encoding.Default.GetString(decryptdata);
            }
        }
        //EDS加密
        private string EDS(string val)
        {
            DESCryptoServiceProvider DesCSP = new DESCryptoServiceProvider();
            MemoryStream ms = new MemoryStream();//先创建 一个内存流
            CryptoStream cryStream = new CryptoStream(ms, DesCSP.CreateEncryptor(), CryptoStreamMode.Write);//将内存流连接到加密转换流
            StreamWriter sw = new StreamWriter(cryStream);
            sw.WriteLine(val);//将要加密的字符串写入加密转换流
            sw.Close();
            cryStream.Close();
            var buffer = ms.ToArray();//将加密后的流转换为字节数组
            return Convert.ToBase64String(buffer);//将加密后的字节数组转换为字符串
        }
        //EDS解密
        private string EDSJ(string val)
        {
            DESCryptoServiceProvider DesCSP = new DESCryptoServiceProvider();
            byte[] encryptdata = Convert.FromBase64String(val);
            MemoryStream ms = new MemoryStream(encryptdata);//将加密后的字节数据加入内存流中
            CryptoStream cryStream = new CryptoStream(ms, DesCSP.CreateDecryptor(), CryptoStreamMode.Read);//内存流连接到解密流中
            StreamReader sr = new StreamReader(cryStream);
            var str= sr.ReadLine();//将解密流读取为字符串
            sr.Close();
            cryStream.Close();
            ms.Close();
            return str;
        }
    }
}
