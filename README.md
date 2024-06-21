# MemListMgr
## Memory List Manager v2.0
#### by Jon Johnson (fafalone), (c) 2024, MIT License

![image](https://github.com/fafalone/MemListMgr/assets/7834493/db57b792-9823-4616-9b55-dae4d9d869ac)

Memory List Manager is a small, lightweight utility to clear standby memory and flush caches, without needing to load a full heavyweight system resources app like SystemInformer, the code for which this project is based on. Windows keeps memory contents it thinks you might use again soon loaded and marked as standby, theoretically this memory is 'available' but memory management isn't always perfect, and high persistent standby and especially modified memory can cause reduced system performance, certain issues with some memory intensive apps or even lead to spurious out of memory errors and system destabilization.

### Updates
Version 2.0 (21 Jun 2024)\
-Added graphical Memory Bar to visualize modified/standby usage\
    Can change colors, double click key (small circle) to set\
-Added Auto-optimize option for common procedure\
-Added option to monitor status and auto-exec AutoOptimize\
-Added minimize to tray option with commands available from a popup menu when right-clicking the tray icon. (Vista+)\
-New settings are saved to/loaded from registry\
-Moved strings to constant list at top for easier translation.

Version 1.2: Display bugs in system memory free and commit charge % labels.\
Version 1.1: Bug in wait cursor code.

### Requirements 
-Windows XP or newer; some features require Windows 7 or Windows 8.1.\
-Program must be run as administrator as it needs the SeProfileSingleProcessPrivilege for basic functions and SeIncreaseQuotaPrivilege for some. 

Build:\
-[twinBASIC Beta 560 or newer](https://github.com/twinbasic/twinbasic/releases) (Note: The manifest specifies requireAdministrator so running the built exe from the IDE will fail if tB is not also running as admin.)\
-(Included in project file) Windows Development Library for twinBASIC v8.3.428 or newer.
Note: To run from the IDE, the project must be built, and built again after any compiler restart. This is due to resource icons only being available from the compiled exe; the app loads them from the exe when running from the IDE.

### Available commands:

**Clear standby RAM:** Clears standby memory of all priority levels.

**Clear Low-Priority Standby RAM Only:** Clear only the low priority standby memory. On Windows 10, this seems to mean Priority 0 only.

**Flush modified RAM:** Commits memory waiting to be written to disk.

**Empty Working Sets:** Empties system and user working set set memory to modified or standby. If you're going to use this, you should use it prior to clearing standby memory.

Extra:

**Combine pages:** De-duplicates certain memory contents.

**Flush registry cache:** Commits pending registry operations to disk.

**Flush system file cache:** Clears files being held in memory for fast access.

### Command line, Jump list, and System Tray

In addition to the GUI, the above commands can be silently run without opening the app GUI by using the following command line switches:

`/clearstdby` - Clear standby memory\
`/clearlpstdby` - Clear low-priority standby memory\
`/flushmod` - Flush modified memory\
`/emptyws` - Empty working sets.\
`/combine` - Combine pages.\
`/flushreg` - Flush registry cache\
`/flushfiles` - Flush file cache

**Note:** Currently, only one command at a time is supported.

On Windows 7 and above, the taskbar icon has a jump list to quickly access commands:

![image](https://github.com/fafalone/MemListMgr/assets/7834493/e7959ccf-679c-44f2-bbae-32f6b8831de5)

On Windows Vista and above, there's an option to minimize to the system tray instead. Left-clicking brings back the main GUI, and right clicking allows access to the commands and exit through a popup menu:

![image](https://github.com/fafalone/MemListMgr/assets/7834493/7779d4c8-c82a-4cea-b3f2-327a48dda285)


### How it works

These actions are done using the `NtSetSystemInformation` API; it's just straightforward calls:

```vba
    Public Function ClearStandby(Optional bLowPriority As Boolean = False) As NTSTATUS
        Dim nCmd As SYSTEM_MEMORY_LIST_COMMAND
        If bLowPriority Then
            nCmd = MemoryPurgeLowPriorityStandbyList
        Else
            nCmd = MemoryPurgeStandbyList
        End If
        Dim status As NTSTATUS
        SetCursor LoadCursor(0, IDC_WAIT)
        status = NtSetSystemInformation(SystemMemoryListInformation, nCmd, LenB(nCmd))
        SetCursor LoadCursor(0, IDC_ARROW)
        Return status
    End Function

    Public Function CombinePages(Optional pNumCombined As LongLong) As NTSTATUS
        Dim mci As MEMORY_COMBINE_INFORMATION_EX
        Dim status As NTSTATUS
        etCursor LoadCursor(0, IDC_WAIT)
        status = NtSetSystemInformation(SystemCombinePhysicalMemoryInformation, mci, LenB(Of MEMORY_COMBINE_INFORMATION_EX))
        pNumCombined = mci.PagesCombined
        SetCursor LoadCursor(0, IDC_ARROW)
        Return status
    End Function

    Public Function FlushRegistryCache() As NTSTATUS
        etCursor LoadCursor(0, IDC_WAIT)
        FlushRegistryCache = NtSetSystemInformation(SystemRegistryReconciliationInformation, ByVal 0&, 0&)
        SetCursor LoadCursor(0, IDC_ARROW)
    End Function

    Public Function FlushFileCache(Optional pSize As LongLong) As NTSTATUS
        Dim sfi As SYSTEM_FILECACHE_INFORMATION
        Dim sfiSet As SYSTEM_FILECACHE_INFORMATION
        Dim status As NTSTATUS
        Dim bRet As Byte
        Dim cb As Long
        etCursor LoadCursor(0, IDC_WAIT)
        status = RtlAdjustPrivilege(SE_INCREASE_QUOTA_PRIVILEGE, 1, 0, bRet)
        status = NtQuerySystemInformation(SystemFileCacheInformationEx, sfi, LenB(Of SYSTEM_FILECACHE_INFORMATION), cb)
        If NT_SUCCESS(status) Then
            sfiSet.MinimumWorkingSet = MAXSIZE_T
            sfiSet.MaximumWorkingSet = MAXSIZE_T
            status = NtSetSystemInformation(SystemFileCacheInformationEx, sfiSet, LenB(Of SYSTEM_FILECACHE_INFORMATION))
            pSize = sfi.CurrentSize
        End If
        SetCursor LoadCursor(0, IDC_ARROW)
        Return status
    End Function
```

That's pretty much all there is to it.

### Made using twinBASIC
![image](https://github.com/fafalone/MemListMgr/assets/7834493/abba1d5d-0adb-4b32-a0cb-0f43ad13f1e6)

twinBASIC is a new language and IDE that's aiming for 100% backwards compatibility with Visual Basic 6.0 while adding a lengthy list of new language features and a modern IDE experience. It's currently under development but far enough along that it supports many large, complex apps. I'm a huge fan of the project and this app was made entirely using only tB.

[twinBASIC FAQs](https://github.com/twinbasic/documentation/wiki/twinBASIC-Frequently-Asked-Questions-(FAQs))

[twinBASIC Home Page](https://twinbasic.com)

[twinBASIC GitHub](https://github.com/twinbasic/twinbasic)

[twinBASIC Releases](https://github.com/twinbasic/twinbasic/releases)
