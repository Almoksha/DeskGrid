using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace DeskGrid.Services;

/// <summary>
/// Shell file operations using IFileOperation COM interface.
/// This provides the same file handling that Windows Explorer uses,
/// handling protected desktop files, shell notifications, and providing
/// proper undo support.
/// </summary>
public class ShellFileOperations : IDisposable
{
    private bool _disposed;

    #region COM Interfaces and Declarations

    // FileOperation CoClass
    [ComImport]
    [Guid("3ad05575-8857-4850-9277-11b85bdb8e09")]
    [ClassInterface(ClassInterfaceType.None)]
    private class FileOperation { }

    // IFileOperation interface
    [ComImport]
    [Guid("947aab5f-0a5c-4c13-b4d6-4bf7836fc9f8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOperation
    {
        uint Advise(IntPtr pfops);
        void Unadvise(uint dwCookie);
        void SetOperationFlags(FileOperationFlags dwOperationFlags);
        void SetProgressMessage([MarshalAs(UnmanagedType.LPWStr)] string pszMessage);
        void SetProgressDialog(IntPtr popd);
        void SetProperties(IntPtr pproparray);
        void SetOwnerWindow(IntPtr hwndOwner);
        void ApplyPropertiesToItem(IShellItem psiItem);
        void ApplyPropertiesToItems([MarshalAs(UnmanagedType.Interface)] object punkItems);
        void RenameItem(IShellItem psiItem, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName, IntPtr pfopsItem);
        void RenameItems([MarshalAs(UnmanagedType.Interface)] object punkItems, [MarshalAs(UnmanagedType.LPWStr)] string pszNewName);
        void MoveItem(IShellItem psiItem, IShellItem psiDestinationFolder, [MarshalAs(UnmanagedType.LPWStr)] string? pszNewName, IntPtr pfopsItem);
        void MoveItems([MarshalAs(UnmanagedType.Interface)] object punkItems, IShellItem psiDestinationFolder);
        void CopyItem(IShellItem psiItem, IShellItem psiDestinationFolder, [MarshalAs(UnmanagedType.LPWStr)] string? pszNewName, IntPtr pfopsItem);
        void CopyItems([MarshalAs(UnmanagedType.Interface)] object punkItems, IShellItem psiDestinationFolder);
        void DeleteItem(IShellItem psiItem, IntPtr pfopsItem);
        void DeleteItems([MarshalAs(UnmanagedType.Interface)] object punkItems);
        uint NewItem(IShellItem psiDestinationFolder, uint dwFileAttributes, [MarshalAs(UnmanagedType.LPWStr)] string pszName, [MarshalAs(UnmanagedType.LPWStr)] string? pszTemplateName, IntPtr pfopsItem);
        void PerformOperations();
        [return: MarshalAs(UnmanagedType.Bool)]
        bool GetAnyOperationsAborted();
    }

    // IShellItem interface
    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    private enum SIGDN : uint
    {
        SIGDN_NORMALDISPLAY = 0x00000000,
        SIGDN_PARENTRELATIVEPARSING = 0x80018001,
        SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000,
        SIGDN_PARENTRELATIVEEDITING = 0x80031001,
        SIGDN_DESKTOPABSOLUTEEDITING = 0x8004c000,
        SIGDN_FILESYSPATH = 0x80058000,
        SIGDN_URL = 0x80068000,
        SIGDN_PARENTRELATIVEFORADDRESSBAR = 0x8007c001,
        SIGDN_PARENTRELATIVE = 0x80080001
    }

    [Flags]
    private enum FileOperationFlags : uint
    {
        FOF_MULTIDESTFILES = 0x0001,
        FOF_CONFIRMMOUSE = 0x0002,
        FOF_SILENT = 0x0004,           // Don't show progress dialog
        FOF_RENAMEONCOLLISION = 0x0008, // Auto-rename if exists
        FOF_NOCONFIRMATION = 0x0010,   // No confirmation dialogs
        FOF_WANTMAPPINGHANDLE = 0x0020,
        FOF_ALLOWUNDO = 0x0040,        // Enable undo (send to recycle bin for delete)
        FOF_FILESONLY = 0x0080,
        FOF_SIMPLEPROGRESS = 0x0100,
        FOF_NOCONFIRMMKDIR = 0x0200,
        FOF_NOERRORUI = 0x0400,        // No error dialogs
        FOF_NOCOPYSECURITYATTRIBS = 0x0800,
        FOF_NORECURSION = 0x1000,
        FOF_NO_CONNECTED_ELEMENTS = 0x2000,
        FOF_WANTNUKEWARNING = 0x4000,
        FOF_NORECURSEREPARSE = 0x8000,
        FOFX_NOSKIPJUNCTIONS = 0x00010000,
        FOFX_PREFERHARDLINK = 0x00020000,
        FOFX_SHOWELEVATIONPROMPT = 0x00040000,
        FOFX_RECYCLEONDELETE = 0x00080000,
        FOFX_EITHERNEGATIVEPROMPTS = 0x00100000,
        FOFX_NOMINIMIZEBOX = 0x01000000,
        FOFX_NOMINIMIZETOAPPTRAY = 0x02000000
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int SHCreateItemFromParsingName(
        [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IntPtr pbc,
        ref Guid riid,
        out IShellItem ppv);

    #endregion

    /// <summary>
    /// Moves a file using Windows Shell (same as Explorer) - silent, no UI
    /// </summary>
    public bool MoveFile(string sourcePath, string destPath)
    {
        return MoveFile(sourcePath, destPath, silent: true, allowUndo: true);
    }

    /// <summary>
    /// Moves a file using Windows Shell
    /// </summary>
    public bool MoveFile(string sourcePath, string destPath, bool silent, bool allowUndo)
    {
        try
        {
            var destFolder = System.IO.Path.GetDirectoryName(destPath);
            var destName = System.IO.Path.GetFileName(destPath);

            if (string.IsNullOrEmpty(destFolder))
                return false;

            // Create FileOperation COM object
            var fileOp = (IFileOperation)new FileOperation();

            // Set flags
            var flags = FileOperationFlags.FOF_NOCONFIRMATION | FileOperationFlags.FOF_NOERRORUI;
            if (silent)
                flags |= FileOperationFlags.FOF_SILENT;
            if (allowUndo)
                flags |= FileOperationFlags.FOF_ALLOWUNDO;

            fileOp.SetOperationFlags(flags);

            // Create shell items
            var sourceItem = CreateShellItem(sourcePath);
            var destFolderItem = CreateShellItem(destFolder);

            if (sourceItem == null || destFolderItem == null)
            {
                System.Diagnostics.Debug.WriteLine("[ShellFileOps] Failed to create shell items");
                return false;
            }

            // Queue the move operation
            fileOp.MoveItem(sourceItem, destFolderItem, destName, IntPtr.Zero);

            // Execute
            fileOp.PerformOperations();

            var aborted = fileOp.GetAnyOperationsAborted();
            System.Diagnostics.Debug.WriteLine($"[ShellFileOps] Move completed, aborted: {aborted}");

            return !aborted;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ShellFileOps] MoveFile failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Copies a file using Windows Shell
    /// </summary>
    public bool CopyFile(string sourcePath, string destPath, bool silent = true)
    {
        try
        {
            var destFolder = System.IO.Path.GetDirectoryName(destPath);
            var destName = System.IO.Path.GetFileName(destPath);

            if (string.IsNullOrEmpty(destFolder))
                return false;

            var fileOp = (IFileOperation)new FileOperation();

            var flags = FileOperationFlags.FOF_NOCONFIRMATION | FileOperationFlags.FOF_NOERRORUI;
            if (silent)
                flags |= FileOperationFlags.FOF_SILENT;

            fileOp.SetOperationFlags(flags);

            var sourceItem = CreateShellItem(sourcePath);
            var destFolderItem = CreateShellItem(destFolder);

            if (sourceItem == null || destFolderItem == null)
                return false;

            fileOp.CopyItem(sourceItem, destFolderItem, destName, IntPtr.Zero);
            fileOp.PerformOperations();

            return !fileOp.GetAnyOperationsAborted();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ShellFileOps] CopyFile failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Deletes a file using Windows Shell (sends to Recycle Bin by default)
    /// </summary>
    public bool DeleteFile(string path, bool useRecycleBin = true)
    {
        try
        {
            var fileOp = (IFileOperation)new FileOperation();

            var flags = FileOperationFlags.FOF_NOCONFIRMATION | FileOperationFlags.FOF_NOERRORUI | FileOperationFlags.FOF_SILENT;
            if (useRecycleBin)
                flags |= FileOperationFlags.FOF_ALLOWUNDO;

            fileOp.SetOperationFlags(flags);

            var item = CreateShellItem(path);
            if (item == null)
                return false;

            fileOp.DeleteItem(item, IntPtr.Zero);
            fileOp.PerformOperations();

            return !fileOp.GetAnyOperationsAborted();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ShellFileOps] DeleteFile failed: {ex.Message}");
            return false;
        }
    }

    private IShellItem? CreateShellItem(string path)
    {
        try
        {
            Guid shellItemGuid = new Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe");
            int hr = SHCreateItemFromParsingName(path, IntPtr.Zero, ref shellItemGuid, out IShellItem item);
            
            if (hr != 0)
            {
                System.Diagnostics.Debug.WriteLine($"[ShellFileOps] SHCreateItemFromParsingName failed for '{path}': 0x{hr:X8}");
                return null;
            }
            
            return item;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ShellFileOps] CreateShellItem failed: {ex.Message}");
            return null;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        // COM objects are automatically released by the runtime
    }
}
