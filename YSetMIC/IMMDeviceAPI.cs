using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace YSetMIC
{
    #region Struct
    public struct WAVEFORMATEX
    {
        ushort wFormatTag;         /* format type */
        ushort nChannels;          /* number of channels (i.e. mono, stereo...) */
        uint nSamplesPerSec;     /* sample rate */
        uint nAvgBytesPerSec;    /* for buffer estimation */
        ushort nBlockAlign;        /* block size of data */
        ushort wBitsPerSample;     /* number of bits per sample of mono data */
        ushort cbSize;             /* the count in bytes of the size of */
        /* extra information (after cbSize) */
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;
    }
    public struct BLOB
    {
        public uint cbSize;
        /* [size_is] */
        public IntPtr/* BYTE* */ pBlobData;
    }


    [StructLayout(LayoutKind.Explicit)]
    public struct PROPVARIANT
    {
        /// <summary>
        /// Value type tag.
        /// </summary>
        [FieldOffset(0)] public short vt;
        /// <summary>
        /// Reserved1.
        /// </summary>
        [FieldOffset(2)] public short wReserved1;
        /// <summary>
        /// Reserved2.
        /// </summary>
        [FieldOffset(4)] public short wReserved2;
        /// <summary>
        /// Reserved3.
        /// </summary>
        [FieldOffset(6)] public short wReserved3;
        /// <summary>
        /// cVal.
        /// </summary>
        [FieldOffset(8)] public sbyte cVal;
        /// <summary>
        /// bVal.
        /// </summary>
        [FieldOffset(8)] public byte bVal;
        /// <summary>
        /// iVal.
        /// </summary>
        [FieldOffset(8)] public short iVal;
        /// <summary>
        /// uiVal.
        /// </summary>
        [FieldOffset(8)] public ushort uiVal;
        /// <summary>
        /// lVal.
        /// </summary>
        [FieldOffset(8)] public int lVal;
        /// <summary>
        /// ulVal.
        /// </summary>
        [FieldOffset(8)] public uint ulVal;
        /// <summary>
        /// intVal.
        /// </summary>
        [FieldOffset(8)] public int intVal;
        /// <summary>
        /// uintVal.
        /// </summary>
        [FieldOffset(8)] public uint uintVal;
        /// <summary>
        /// hVal.
        /// </summary>
        [FieldOffset(8)] public long hVal;
        /// <summary>
        /// uhVal.
        /// </summary>
        [FieldOffset(8)] public long uhVal;
        /// <summary>
        /// fltVal.
        /// </summary>
        [FieldOffset(8)] public float fltVal;
        /// <summary>
        /// dblVal.
        /// </summary>
        [FieldOffset(8)] public double dblVal;
        //VARIANT_BOOL boolVal;
        /// <summary>
        /// boolVal.
        /// </summary>
        [FieldOffset(8)] public short boolVal;
        /// <summary>
        /// scode.
        /// </summary>
        [FieldOffset(8)] public int scode;
        //CY cyVal;
        //[FieldOffset(8)] private DateTime date; - can cause issues with invalid value
        /// <summary>
        /// Date time.
        /// </summary>
        [FieldOffset(8)] public System.Runtime.InteropServices.ComTypes.FILETIME filetime;
        //CLSID* puuid;
        //CLIPDATA* pclipdata;
        //BSTR bstrVal;
        //BSTRBLOB bstrblobVal;
        /// <summary>
        /// Binary large object.
        /// </summary>
        [FieldOffset(8)] public BLOB blobVal;
        //LPSTR pszVal;
        /// <summary>
        /// Pointer value.
        /// </summary>
        [FieldOffset(8)] public IntPtr pointerValue; //LPWSTR 
                                                     //IUnknown* punkVal;
        /*IDispatch* pdispVal;
        IStream* pStream;
        IStorage* pStorage;
        LPVERSIONEDSTREAM pVersionedStream;
        LPSAFEARRAY parray;
        CAC cac;
        CAUB caub;
        CAI cai;
        CAUI caui;
        CAL cal;
        CAUL caul;
        CAH cah;
        CAUH cauh;
        CAFLT caflt;
        CADBL cadbl;
        CABOOL cabool;
        CASCODE cascode;
        CACY cacy;
        CADATE cadate;
        CAFILETIME cafiletime;
        CACLSID cauuid;
        CACLIPDATA caclipdata;
        CABSTR cabstr;
        CABSTRBLOB cabstrblob;
        CALPSTR calpstr;
        CALPWSTR calpwstr;
        CAPROPVARIANT capropvar;
        CHAR* pcVal;
        UCHAR* pbVal;
        SHORT* piVal;
        USHORT* puiVal;
        LONG* plVal;
        ULONG* pulVal;
        INT* pintVal;
        UINT* puintVal;
        FLOAT* pfltVal;
        DOUBLE* pdblVal;
        VARIANT_BOOL* pboolVal;
        DECIMAL* pdecVal;
        SCODE* pscode;
        CY* pcyVal;
        DATE* pdate;
        BSTR* pbstrVal;
        IUnknown** ppunkVal;
        IDispatch** ppdispVal;
        LPSAFEARRAY* pparray;
        PROPVARIANT* pvarVal;
        */

        /// <summary>
        /// Creates a new PropVariant containing a long value
        /// </summary>
        public static PROPVARIANT FromLong(long value)
        {
            return new PROPVARIANT() { vt = (short)VarEnum.VT_I8, hVal = value };
        }

        /// <summary>
        /// Helper method to gets blob data
        /// </summary>
        private byte[] GetBlob()
        {
            var blob = new byte[blobVal.cbSize];
            Marshal.Copy(blobVal.pBlobData, blob, 0, blob.Length);
            return blob;
        }

        /// <summary>
        /// Interprets a blob as an array of structs
        /// </summary>
        public T[] GetBlobAsArrayOf<T>()
        {
            var blobByteLength = blobVal.cbSize;
            var singleInstance = (T)Activator.CreateInstance(typeof(T));
            var structSize = Marshal.SizeOf(singleInstance);
            if (blobByteLength % structSize != 0)
            {
                throw new Exception(String.Format("Blob size {0} not a multiple of struct size {1}", blobByteLength, structSize));
            }
            var items = blobByteLength / structSize;
            var array = new T[items];
            for (int n = 0; n < items; n++)
            {
                array[n] = (T)Activator.CreateInstance(typeof(T));
                Marshal.PtrToStructure(new IntPtr((long)blobVal.pBlobData + n * structSize), array[n]);
            }
            return array;
        }

        /// <summary>
        /// Gets the type of data in this PropVariant
        /// </summary>
        public VarEnum DataType => (VarEnum)vt;

        /// <summary>
        /// Property value
        /// </summary>
        public object Value
        {
            get
            {
                VarEnum ve = DataType;
                switch (ve)
                {
                    case VarEnum.VT_I1:
                        return bVal;
                    case VarEnum.VT_I2:
                        return iVal;
                    case VarEnum.VT_I4:
                        return lVal;
                    case VarEnum.VT_I8:
                        return hVal;
                    case VarEnum.VT_INT:
                        return iVal;
                    case VarEnum.VT_UI4:
                        return ulVal;
                    case VarEnum.VT_UI8:
                        return uhVal;
                    case VarEnum.VT_LPWSTR:
                        return Marshal.PtrToStringUni(pointerValue);
                    case VarEnum.VT_BLOB:
                    case VarEnum.VT_VECTOR | VarEnum.VT_UI1:
                        return GetBlob();
                    case VarEnum.VT_CLSID:
                        return Marshal.PtrToStructure<Guid>(pointerValue);
                    case VarEnum.VT_BOOL:
                        switch (boolVal)
                        {
                            case -1:
                                return true;
                            case 0:
                                return false;
                            default:
                                throw new NotSupportedException("PropVariant VT_BOOL must be either -1 or 0");
                        }
                    case VarEnum.VT_FILETIME:
                        return DateTime.FromFileTime((((long)filetime.dwHighDateTime) << 32) + filetime.dwLowDateTime);
                }
                throw new NotImplementedException("PropVariant " + ve);
            }
        }

        /// <summary>
        /// Clears with a known pointer
        /// </summary>
        public static void Clear(IntPtr ptr)
        {
            Marshal.FreeHGlobal(ptr);
        }
    }


    #endregion
    #region Enum
    public enum EDataFlow
    {
        eRender = 0,
        eCapture = (eRender + 1),
        eAll = (eCapture + 1),
        EDataFlow_enum_count = (eAll + 1)
    }

    public enum StorageAccessMode
    {
        /// <summary>
        /// Read-only access mode.
        /// </summary>
        Read,
        /// <summary>
        /// Write-only access mode.
        /// </summary>
        Write,
        /// <summary>
        /// Read-write access mode.
        /// </summary>
        ReadWrite
    }
    [Flags]
    public enum DeviceState
    {
        /// <summary>
        /// DEVICE_STATE_ACTIVE
        /// </summary>
        Active = 0x00000001,
        /// <summary>
        /// DEVICE_STATE_DISABLED
        /// </summary>
        Disabled = 0x00000002,
        /// <summary>
        /// DEVICE_STATE_NOTPRESENT 
        /// </summary>
        NotPresent = 0x00000004,
        /// <summary>
        /// DEVICE_STATE_UNPLUGGED
        /// </summary>
        Unplugged = 0x00000008,
        /// <summary>
        /// DEVICE_STATEMASK_ALL
        /// </summary>
        All = 0x0000000F
    }

    public enum ERole
    {
        eConsole = 0,
        eMultimedia = (eConsole + 1),
        eCommunications = (eMultimedia + 1),
        ERole_enum_count = (eCommunications + 1)
    }
    #endregion
    #region Com Interface
    [Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore
    {

        int GetCount([Out] out uint cProps);

        int GetAt([In] uint iProp, [Out] out PROPERTYKEY pkey);

        int GetValue([In] ref PROPERTYKEY key, [Out] out PROPVARIANT pv);
        int SetValue([In] ref PROPERTYKEY key, [In] ref PROPVARIANT propvar);

        int Commit();

    };
    [Guid("7991EEC9-7E89-4D85-8390-6C703CEC60C0"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMNotificationClient
    {
        /// <summary>
        /// Device State Changed
        /// </summary>
        void OnDeviceStateChanged([MarshalAs(UnmanagedType.LPWStr)] string deviceId, [MarshalAs(UnmanagedType.I4)] uint newState);

        /// <summary>
        /// Device Added
        /// </summary>
        void OnDeviceAdded([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId);

        /// <summary>
        /// Device Removed
        /// </summary>
        void OnDeviceRemoved([MarshalAs(UnmanagedType.LPWStr)] string deviceId);

        /// <summary>
        /// Default Device Changed
        /// </summary>
        void OnDefaultDeviceChanged(EDataFlow flow, ERole role, [MarshalAs(UnmanagedType.LPWStr)] string defaultDeviceId);

        /// <summary>
        /// Property Value Changed
        /// </summary>
        /// <param name="pwstrDeviceId"></param>
        /// <param name="key"></param>
        void OnPropertyValueChanged([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId, PROPERTYKEY key);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    interface IMMDevice
    {
        int Activate([In] Guid iid, uint dwClsCtx, [In] ref IntPtr pActivationParams, [Out][MarshalAs(UnmanagedType.IUnknown)] out IntPtr ppInterface);

        int OpenPropertyStore(StorageAccessMode stgmAccess, out IPropertyStore ppProperties);

        int GetId([Out][MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);

        int GetState([Out] out uint pdwState);
    }
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
    interface IMMDeviceCollection
    {
        int GetCount(out uint pcDevices);

        int Item([In] uint nDevice, [Out] out IMMDevice ppDevice);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [CoClass(typeof(MMDeviceEnumerator))]
    interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(EDataFlow dataFlow, [In] DeviceState stateMask, out IMMDeviceCollection devices);


        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice endpoint);

        int GetDevice(string id, out IMMDevice deviceName);

        int RegisterEndpointNotificationCallback(IMMNotificationClient client);

        int UnregisterEndpointNotificationCallback(IMMNotificationClient client);
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    class MMDeviceEnumerator { }


    [ComImport]
    [Guid("f8679f50-850a-41cf-9c72-430f290290c8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [CoClass(typeof(CPolicyConfigClient))]
    interface IPolicyConfig
    {
        int GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string context, ref IntPtr _waveformatex);
        int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string context, int i, ref IntPtr _waveformatex);
        int ResetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string context);

        int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string context, ref WAVEFORMATEX _waveformatex1, ref WAVEFORMATEX _waveformatex2);

        int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string context, int i, ref long j, ref long k);

        int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string context, ref long i);

        int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string context, IntPtr deviceShareMode);

        int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string context, IntPtr deviceShareMode);

        int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string context, ref PROPERTYKEY propertyKey1, ref PROPERTYKEY propertyVariant);

        int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string context, ref PROPERTYKEY propertyKey1, ref PROPERTYKEY propertyVariant);

        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string wszDeviceId, ERole eRole);

        int SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string context, int i);
    }
    [ComImport]
    [Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
    class CPolicyConfigClient
    {

    }

    #endregion
}
