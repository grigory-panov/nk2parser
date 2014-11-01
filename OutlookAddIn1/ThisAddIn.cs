using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.IO;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {

        const ushort MV_FLAG = 0x1000;
        const ushort PT_UNSPECIFIED = 0x0;
        const ushort PT_NULL = 0x1;
        const ushort PT_I2 = 0x2;
        const ushort PT_SHORT = 0x2;
        const ushort PT_LONG = 0x3;
        const ushort PT_FLOAT = 0x4;
        const ushort PT_DOUBLE = 0x5;
        const ushort PT_CURRENCY = 0x6;
        const ushort PT_APPTIME = 0x7;
        const ushort PT_ERROR = 0xa;
        const ushort PT_BOOLEAN = 0xb;
        const ushort PT_OBJECT = 0xd;
        const ushort PT_I8 = 0x14;
        const ushort PT_STRING8 = 0x1e;
        const ushort PT_UNICODE = 0x1f;
        const ushort PT_SYSTIME = 0x40;
        const ushort PT_CLSID = 0x48;
        const ushort PT_SVREID = 0xFB;
        const ushort PT_SRESTRICT = 0xFD;
        const ushort PT_ACTIONS = 0xFE;
        const ushort PT_BINARY = 0x102;
        /* Multi valued property types */
        const ushort PT_MV_SHORT = (MV_FLAG | PT_SHORT);
        const ushort PT_MV_LONG = (MV_FLAG | PT_LONG);
        const ushort PT_MV_FLOAT = (MV_FLAG | PT_FLOAT);
        const ushort PT_MV_DOUBLE = (MV_FLAG | PT_DOUBLE);
        const ushort PT_MV_CURRENCY = (MV_FLAG | PT_CURRENCY);
        const ushort PT_MV_APPTIME = (MV_FLAG | PT_APPTIME);
        const ushort PT_MV_I8 = (MV_FLAG | PT_I8);
        const ushort PT_MV_STRING8 = (MV_FLAG | PT_STRING8);
        const ushort PT_MV_UNICODE = (MV_FLAG | PT_UNICODE);
        const ushort PT_MV_SYSTIME = (MV_FLAG | PT_SYSTIME);
        const ushort PT_MV_CLSID = (MV_FLAG | PT_CLSID);
        const ushort PT_MV_BINARY = (MV_FLAG | PT_BINARY);
        const uint PR_ENTRYID = 0x0FFF0102;
        const uint PR_DISPLAY_NAME_W = 0x3001001F;
        const uint PR_DROPDOWN_DISPLAY_NAME_W = 0x6003001F;
        const uint PR_EMAIL_ADDRESS_W = 0x3003001F;
        private string _state;
        private byte[] _buffer = new byte[12];
        System.Collections.ArrayList list = new System.Collections.ArrayList();
        private const string ERROR = "ERROR";
        private const string METADATA_BEGIN = "METADATA_BEGIN";
        private const string METADATA_END = "METADATA_END";
        private const string NUMBER_OF_ROWS = "NUMBER_OF_ROWS";
        private const string NUMBER_OF_PROPERTIES = "NUMBER_OF_PROPERTIES";
        private const string PROPERTY_TAG = "PROPERTY_TAG";
        private const string PROPERTY_RESERVED_DATA = "PROPERTY_RESERVED_DATA";
        private const string PROPERTY_VALUE_UNION = "PROPERTY_VALUE_UNION";
        private const string PROPERTY_VALUE_DATA = "PROPERTY_VALUE_DATA";
        private const string NUMBER_OF_ARRAY_ELEMENTS = "NUMBER_OF_ARRAY_ELEMENTS";
        private const string NUMBER_OF_VALUE_BYTES = "NUMBER_OF_VALUE_BYTES";


        public string state
        {
            get { return _state; }
            set
            {
                Trace.TraceInformation(value);
                _state = value;
            }
        }
        public byte[] buffer
        {
            get
            {
                return _buffer;
            }
            set
            {
                // сохраняем старое значение
                list.AddRange(_buffer);
                _buffer = value;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.OperatingSystem osInfo = System.Environment.OSVersion;
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            //if (osInfo.Version.Major <= 5)
            //{
            filePath += "\\Microsoft\\Outlook\\Outlook.NK2";
            FileStream fs = null;
            try
            {
                fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

                byte[] content = loadContent(fs);
                if (content == null)
                {
                    Trace.TraceInformation("File NK2 not found, exit.");
                    return;
                }
                if (content.Length == 0)
                {
                    Trace.TraceInformation("File NK2 is empty, exit.");
                    return;
                }
                // в Outlook 2010-2013 лист с автодополнением хранится в inbox в скрытом письме с темой IPM.Configuration.Autocomplete
                // в 2003-2007 он лежит как файл формата .nk2
                Application oApp = Globals.ThisAddIn.Application;
                MAPIFolder inboxFolder = oApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                MAPIFolder contactsFolder = oApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                //string filePath = "C:\\Users\\User\\Documents\\Professional 2013\\Outlook.NK2";
                //string filePathModified = "C:\\Users\\User\\Documents\\Professional 2013\\Outlook.NK2";
                //FileStream fileStream = File.OpenRead(filePath);
                int replacementPosition = 0;
                int replacementPosition2 = 0;

                //byte[] content = new byte[fileStream.Length];

                /*int numBytesToRead = (int)fileStream.Length;
                int numBytesRead = 0;
                while (numBytesToRead > 0)
                {
                    int n = fileStream.Read(content, numBytesRead, numBytesToRead);
                    if (n == 0)
                        break;
                    numBytesRead += n;
                    numBytesToRead -= n;
                }*/
                int idx = 0;
                uint propertyValueDataLength = 0;
                uint propertiesCount = 0;
                uint arrayElementsCount = 0;
                bool inArray = false;
                uint rowsCount = 0;
                uint currentTag = 0;
                string textToReplace = null;
                string email = null;
                ushort propertyValueType = PT_UNSPECIFIED;
                state = METADATA_BEGIN;
                int i = 0;
                while (state != ERROR && i < content.Length)
                {
                    switch (state)
                    {
                        case METADATA_BEGIN:
                            if (idx < 12)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                idx = 0;
                                printBinary(buffer, 12);
                                buffer = new byte[4];
                                state = NUMBER_OF_ROWS;
                            }
                            break;
                        case METADATA_END:
                            if (idx < 12)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                state = ERROR;
                            }
                            break;
                        case NUMBER_OF_ROWS:
                            if (idx < 4)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                idx = 0;
                                rowsCount = BitConverter.ToUInt32(buffer, 0);
                                printUInt(buffer);
                                buffer = new byte[4];
                                state = NUMBER_OF_PROPERTIES;
                            }
                            break;
                        case NUMBER_OF_PROPERTIES:
                            if (idx < 4)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                printUInt(buffer);
                                idx = 0;
                                propertiesCount = BitConverter.ToUInt32(buffer, 0);
                                buffer = new byte[4];
                                state = PROPERTY_TAG;
                            }
                            break;
                        case PROPERTY_TAG:
                            if (idx < 4)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                idx = 0;
                                currentTag = BitConverter.ToUInt32(buffer, 0);
                                printTag(buffer);
                                propertyValueType = BitConverter.ToUInt16(buffer, 0);
                                buffer = new byte[4];
                                state = PROPERTY_RESERVED_DATA;
                            }
                            break;
                        case PROPERTY_RESERVED_DATA:
                            if (idx < 4)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                idx = 0;
                                printUInt(buffer);
                                buffer = new byte[8];
                                state = PROPERTY_VALUE_UNION;
                            }
                            break;
                        case PROPERTY_VALUE_UNION:
                            if (idx < 8)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                printULong(buffer);
                                Trace.TraceInformation("-------------------------------------------");
                                arrayElementsCount = 0;
                                inArray = false;
                                if (propertyValueType == PT_BINARY
                                    || propertyValueType == PT_STRING8
                                    || propertyValueType == PT_UNICODE)
                                {
                                    state = NUMBER_OF_VALUE_BYTES;
                                }
                                else if (propertyValueType == PT_MV_BINARY
                                    || propertyValueType == PT_MV_STRING8
                                    || propertyValueType == PT_MV_UNICODE)
                                {
                                    state = NUMBER_OF_ARRAY_ELEMENTS;
                                }
                                else
                                {
                                    state = PROPERTY_VALUE_DATA;
                                }
                                propertyValueDataLength = calculateLengthByType(propertyValueType);
                                idx = 0;
                                buffer = new byte[propertyValueDataLength];
                            }
                            break;
                        case PROPERTY_VALUE_DATA:
                            if (idx < propertyValueDataLength)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                printVarious(buffer, propertyValueType);
                                Trace.TraceInformation("-------------------------------------------");
                                if (inArray && arrayElementsCount > 0)
                                {
                                    arrayElementsCount--;
                                    state = NUMBER_OF_VALUE_BYTES;
                                    idx = 0;
                                    buffer = new byte[4];
                                }
                                else
                                {

                                    if (currentTag == PR_DROPDOWN_DISPLAY_NAME_W)
                                    {
                                        replacementPosition = list.Count;
                                    }
                                    if (currentTag == PR_DISPLAY_NAME_W)
                                    {
                                        replacementPosition2 = list.Count;
                                    }
                                    if (currentTag == PR_EMAIL_ADDRESS_W)
                                    {
                                        email = System.Text.Encoding.Unicode.GetString(buffer);
                                        Trace.TraceInformation("email = " + email);
                                    }
                                    propertiesCount--;
                                    Trace.TraceInformation("END OF PROPERTY, LEFT:" + propertiesCount);
                                    idx = 0;
                                    if (propertiesCount > 0)
                                    {
                                        state = PROPERTY_TAG;
                                        buffer = new byte[4];
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(email))
                                        {
                                            textToReplace = findFullNameInAdressBook(email, contactsFolder);
                                        }
                                        if (replacementPosition > replacementPosition2)
                                        {
                                            if (replacementPosition != -1 && !string.IsNullOrEmpty(textToReplace))
                                            {
                                                replace(replacementPosition - 4, textToReplace);
                                            }
                                            if (replacementPosition2 != -1 && !string.IsNullOrEmpty(textToReplace))
                                            {
                                                replace(replacementPosition2 - 4, textToReplace);
                                            }

                                        }
                                        else
                                        {
                                            if (replacementPosition2 != -1 && !string.IsNullOrEmpty(textToReplace))
                                            {
                                                replace(replacementPosition2 - 4, textToReplace);
                                            }
                                            if (replacementPosition != -1 && !string.IsNullOrEmpty(textToReplace))
                                            {
                                                replace(replacementPosition - 4, textToReplace);
                                            }
                                        }
                                        replacementPosition = -1;
                                        replacementPosition2 = -1;
                                        textToReplace = null;
                                        email = null;
                                        rowsCount--;
                                        Trace.TraceInformation("END OF ROW, LEFT:" + rowsCount);
                                        if (rowsCount > 0)
                                        {
                                            state = NUMBER_OF_PROPERTIES;
                                            buffer = new byte[4];
                                        }
                                        else
                                        {
                                            state = METADATA_END;
                                            buffer = new byte[12];
                                        }
                                    }
                                }
                            }
                            break;
                        case NUMBER_OF_VALUE_BYTES:
                            if (idx < 4)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                propertyValueDataLength = BitConverter.ToUInt32(buffer, 0);
                                printUInt(buffer);
                                if (propertyValueDataLength > 100000)
                                {
                                    Trace.TraceError("too long property value: " + propertyValueDataLength);
                                    state = ERROR;
                                    break;
                                }
                                buffer = new byte[propertyValueDataLength];
                                idx = 0;
                                state = PROPERTY_VALUE_DATA;
                            }
                            break;
                        case NUMBER_OF_ARRAY_ELEMENTS:
                            if (idx < propertyValueDataLength)
                            {
                                buffer[idx] = content[i++];
                                idx++;
                            }
                            else
                            {
                                printUInt(buffer);
                                arrayElementsCount = BitConverter.ToUInt32(buffer, 0);
                                inArray = true;
                                buffer = new byte[4];
                                idx = 0;
                                state = NUMBER_OF_VALUE_BYTES;
                            }
                            break;
                        default:
                            state = ERROR;
                            break;
                    }
                }
                printBinary(buffer, 12);
                buffer = null;
                Trace.TraceInformation("end");
                //fileStream.Close();
                byte[] a = (byte[])list.ToArray(typeof(byte));
                saveContent(a, fs);
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
            //File.WriteAllBytes(filePathModified, a);
        }

        private void saveContent(byte[] a, FileStream fs)
        {
            fs.Position = 0;
            fs.Write(a, 0, a.Length);
            fs.Flush();

        }

        private byte[] loadContent(FileStream fileStream)
        {


            byte[] c = new byte[fileStream.Length];
            int numBytesToRead = (int)fileStream.Length;
            int numBytesRead = 0;
            while (numBytesToRead > 0)
            {
                int n = fileStream.Read(c, numBytesRead, numBytesToRead);
                if (n == 0)
                    break;
                numBytesRead += n;
                numBytesToRead -= n;
            }
            return c;



            /* StorageItem storage = null;
             try
             {
                 storage = inboxFolder.GetStorage("IPM.Configuration.Autocomplete", OlStorageIdentifierType.olIdentifyBySubject);
             }
             catch (System.Exception ex)
             {
                 Console.WriteLine(ex.Message);
                 return;
             }

             PropertyAccessor propertyAcc = storage.PropertyAccessor;
             // получаем кеш как массив байт
             string PR_ROAMING_BINARYSTREAM = "http://schemas.microsoft.com/mapi/proptag/0x7C090102";
             byte[] content = (byte[]) propertyAcc.GetProperty(PR_ROAMING_BINARYSTREAM);// );
             */


        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private string findFullNameInAdressBook(string email, MAPIFolder inbox)
        {

            ContactItem result = (ContactItem)inbox.Items.Find(string.Format("[Email1address]='{0}'", email.TrimEnd('\0')));
            if (result != null)
            {
                if (email.Contains("@"))
                {
                    return string.Format("{0} <{1}>\0", result.FullName, email.TrimEnd('\0'));
                }
                else
                {
                    return string.Format("{0}\0", result.FullName);
                }
            }
            else
            {
                Trace.TraceInformation(email + " not found in address book");
                return null;
            }

        }
        private void replace(int pos1, string textToReplace)
        {
            byte[] len = new byte[4];
            len[0] = (byte)list[pos1];
            len[1] = (byte)list[pos1 + 1];
            len[2] = (byte)list[pos1 + 2];
            len[3] = (byte)list[pos1 + 3];
            uint lenght = BitConverter.ToUInt32(len, 0);

            // Console.WriteLine(BitConverter.ToString((byte[])list.GetRange(pos1, (int)lenght + 4).ToArray(typeof(byte))));

            list.RemoveRange(pos1, (int)lenght + 4);
            byte[] newText = System.Text.Encoding.Unicode.GetBytes(textToReplace);
            uint newLenght = (uint)newText.Length;
            list.InsertRange(pos1, BitConverter.GetBytes(newLenght));
            list.InsertRange(pos1 + 4, newText);

            //Console.WriteLine(BitConverter.ToString((byte[])list.GetRange(pos1, (int)newLenght + 4).ToArray(typeof(byte))));
        }


        private uint calculateLengthByType(ushort type)
        {

            if (type == PT_I2 || type == PT_LONG || type == PT_FLOAT || type == PT_DOUBLE
                || type == PT_BOOLEAN || type == PT_SYSTIME || type == PT_I8 || type == PT_ERROR)
            {
                return 0;
            }
            if (type == PT_CLSID)
            {
                return 16;
            }
            return 4; // в 4-х первых байтах в зависимости от типа лежит либо длина в байтах если это строка либо длина массива, если это массив. мы изменим длину буфера после прочтения ътих 4-х байт

        }

        private void printVarious(byte[] buffer, ushort type)
        {
            if (type == PT_UNICODE)
            {
                Trace.TraceInformation(System.Text.Encoding.Unicode.GetString(buffer));
            }
            else if (type == PT_STRING8)
            {
                Trace.TraceInformation(System.Text.Encoding.ASCII.GetString(buffer));
            }
            else
            {
                Trace.TraceInformation(BitConverter.ToString(buffer, 0, buffer.Length));
            }

        }
        private void printBinary(byte[] buffer, int len)
        {
            // various 
            Trace.TraceInformation(BitConverter.ToString(buffer, 0, len));
        }
        private void printULong(byte[] buffer)
        {
            //8 byte
            ulong value = BitConverter.ToUInt64(buffer, 0);
            Trace.TraceInformation("Hex: {0:x}", value);
        }
        private void printUInt(byte[] buffer)
        {
            //4 byte
            uint value = BitConverter.ToUInt32(buffer, 0);
            Trace.TraceInformation("Hex: {0:x}", value);
        }
        private void printTag(byte[] buffer)
        {
            //4 byte
            uint value = BitConverter.ToUInt32(buffer, 0);

            if (value == PR_DISPLAY_NAME_W)
            {
                Trace.TraceInformation("PR_DISPLAY_NAME_W, Hex: {0:x}", value);
            }
            if (value == PR_DROPDOWN_DISPLAY_NAME_W)
            {
                Trace.TraceInformation("PR_DROPDOWN_DISPLAY_NAME_W, Hex: {0:x}", value);
            }
            else if (value == PR_ENTRYID)
            {
                Trace.TraceInformation("PR_ENTRYID, Hex: {0:x}", value);
            }
            else
            {
                Trace.TraceInformation("Not important tag, Hex: {0:x}", value);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
