# MemListMgr
## Memory List Manager v1.0
#### by Jon Johnson (fafalone), (c) 2024, MIT License

![image](https://github.com/fafalone/MemListMgr/assets/7834493/fa2bf40d-1936-4e17-8a77-f2e85500fc4b)

Memory List Manager is a small, lightweight utility to clear standby memory and flush caches, without needing to load a full heavyweight system resources app like SystemInformer, the code for which this project is based on. Windows keeps memory contents it thinks you might use again soon loaded and marked as standby, theoretically this memory is 'available' but memory management isn't always perfect, and high persistent standby and especially modified memory can cause reduced system performance, certain issues with some memory intensive apps or even lead to spurious out of memory errors and system destabilization.

### Requirements
-Windows XP or newer; some features require Windows 7 or Windows 8.1.\
-Program must be run as administrator as it needs the SeProfileSingleProcessPrivilege for basic functions and SeIncreaseQuotaPrivilege for some.\

Build:\
-[twinBASIC Beta 553 or newer](https://github.com/twinbasic/twinbasic/releases) (Note: The manifest specifies requireAdministrator so running the built exe from the IDE will fail if tB is not also running as admin.)\
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

### Command line and Jump list

In addition to the GUI, the above commands can be silently run without opening the app GUI by using the following command line switches:

`/clearstdby` - Clear standby memory\
`/clearlpstdby` - Clear low-priority standby memory\
`/flushmod` = Flush modified memory\
`/emptyws` - Empty working sets.\
`/combine` - Combine pages.\
`/flushreg` - Flush registry cache\
`/flushfiles` - Flush file cache

**Note:** Currently, only one command at a time is supported.

Further, on Windows 7 and above, the taskbar icon has a jump list to quickly access commands:

![image](https://github.com/fafalone/MemListMgr/assets/7834493/e7959ccf-679c-44f2-bbae-32f6b8831de5)

### Made using twinBASIC
![image](https://github.com/fafalone/MemListMgr/assets/7834493/abba1d5d-0adb-4b32-a0cb-0f43ad13f1e6)

twinBASIC is a new language and IDE that's aiming for 100% backwards compatibility with Visual Basic 6.0 while adding a lengthy list of new language features and a modern IDE experience. It's currently under development but far enough along that it supports many large, complex apps. I'm a huge fan of the project and this app was made entirely using only tB.

[twinBASIC FAQs](https://github.com/twinbasic/documentation/wiki/twinBASIC-Frequently-Asked-Questions-(FAQs))

[twinBASIC Home Page](https://twinbasic.com)

[twinBASIC GitHub](https://github.com/twinbasic/twinbasic)

[twinBASIC Releases](https://github.com/twinbasic/twinbasic/releases)
