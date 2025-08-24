using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.xls.XlsFileFormat;

namespace Macrome
{
    public class RecordHelper
    {
        public static readonly List<RecordType> RelevantTypes = new List<RecordType>()
        {
            RecordType.BoundSheet8, //Sheet definitions (Defines macro sheets + hides them)
            RecordType.Lbl,         //Named Cells (Contains Auto_Start) 
            RecordType.Formula,     //The meat of most cell content
            RecordType.SupBook,     //Contains information for cross-sheet references
            RecordType.ExternSheet, //Contains the XTI records mapping ixti values to BoundSheet8
            RecordType.FilePass,    //Indicates the presence of an RC4 or XOR Obfuscation Password
        };

        public static string GetRelevantRecordDumpString(WorkbookStream wbs, 
            bool dumpHexBytes = false, 
            bool showAttrInfo = false)
        {
            int numBytesToDump = 0;
            if (dumpHexBytes) numBytesToDump = 0x1000;

            bool hasPassword = wbs.HasPasswordToOpen();

            List<BiffRecord> relevantRecords = wbs.Records
                .Where(rec => RecordHelper.RelevantTypes.Contains(rec.Id)
                ).ToList();

            //We can only interpret the data of these records if they are not encrypted
            if (!hasPassword)
            {
                relevantRecords = RecordHelper.ConvertToSpecificRecords(relevantRecords);
                relevantRecords = PtgHelper.UpdateGlobalsStreamReferences(relevantRecords);
            }

            string dumpString = "";

            foreach (var record in relevantRecords)
            {
                try
                {
                    dumpString += record.Sheet.stName + ":";
                    dumpString += record.ToHexDumpString(numBytesToDump, showAttrInfo);
                    dumpString += "\n";
                }
                catch { }
            }

            return dumpString;
        }


        public static List<BiffRecord> ConvertToSpecificRecords(List<BiffRecord> generalRecords)
        {
            List<BiffRecord> specificRecords = new List<BiffRecord>();
            foreach (var record in generalRecords)
            {
                BiffRecord result = GetSpecificRecord(record);
                specificRecords.Add(result);
            }

            return specificRecords;
        }

        public static BiffRecord GetSpecificRecord(BiffRecord record)
        {
            BiffRecord result = null;
            switch (record.Id)
            {
                case RecordType.Formula:
                    result = record.AsRecordType<Formula>();
                    break;
                case RecordType.Lbl:
                    result = record.AsRecordType<Lbl>();
                    break;
                case RecordType.BoundSheet8:
                    result = record.AsRecordType<BoundSheet8>();
                    break;
                case RecordType.SupBook:
                    result = record.AsRecordType<SupBook>();
                    break;
                case RecordType.ExternSheet:
                    result = record.AsRecordType<ExternSheet>();
                    break;
                default:
                    result = record;
                    break;
            }

            return result;
        }

        public static List<BiffRecord> ParseBiffStreamBytes(byte[] bytes)
        {
            List<BiffRecord> records = new List<BiffRecord>();
            MemoryStream ms = new MemoryStream(bytes);
            VirtualStreamReader vsr = new VirtualStreamReader(ms);

            while (vsr.BaseStream.Position < vsr.BaseStream.Length)
            {
                RecordType id = (RecordType)vsr.ReadUInt16();

                if (id == 0)
                {
                    // Console.WriteLine("RecordID == 0 - stopping");
                    break;
                }


                UInt16 length = vsr.ReadUInt16();

                BiffRecord br = new BiffRecord(vsr, id, length);
                
                vsr.ReadBytes(length);
                records.Add(br);
            }

            return records;
        }

        public static byte[] ConvertBiffRecordsToBytes(List<BiffRecord> records)
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);
            foreach (var record in records)
            {
                bw.Write(record.GetBytes());
            }
            return bw.GetBytesWritten();
        }

    }
}
