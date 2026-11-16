// SetLdapABDefault.cpp
// Sets an Outlook address list as the default directory and first entry in
// the address-book search path for the current MAPI profile.
//
// Build (32-bit):
//   i686-w64-mingw32-g++ -O2 -municode -static -static-libgcc -static-libstdc++
//     -o SetLdapABDefault_x86.exe SetLdapABDefault.cpp -lmapi32 -lole32 -loleaut32
// Build (64-bit):
//   x86_64-w64-mingw32-g++ -O2 -municode -static -static-libgcc -static-libstdc++
//     -o SetLdapABDefault_x64.exe SetLdapABDefault.cpp -lmapi32 -lole32 -loleaut32
//
// Note: this intentionally does not import FreeProws. Some Outlook/MAPI
// installations expose a reduced mapi32.dll that does not export it. Row sets
// are released with MAPIFreeBuffer instead.

#define WIN32_LEAN_AND_MEAN
#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <mapix.h>
#include <mapiutil.h>
#include <mapitags.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>

enum {
    RC_OK = 0,
    RC_FAIL = 1,
    RC_USAGE = 2,
    RC_NOT_FOUND = 3
};

static void PrintHR(const char* what, HRESULT hr) {
    wprintf(L"%hs: hr=0x%08lx\n", what, (unsigned long)hr);
}

static void FreeRowSetManually(LPSRowSet rows) {
    if (!rows) return;
    for (ULONG i = 0; i < rows->cRows; ++i) {
        if (rows->aRow[i].lpProps) MAPIFreeBuffer(rows->aRow[i].lpProps);
    }
    MAPIFreeBuffer(rows);
}

static LPSPropValue FindProp(LPSRow row, ULONG propertyId) {
    if (!row) return NULL;
    for (ULONG i = 0; i < row->cValues; ++i) {
        if (PROP_ID(row->lpProps[i].ulPropTag) == propertyId) {
            return &row->lpProps[i];
        }
    }
    return NULL;
}

static const wchar_t* GetDisplayName(LPSRow row) {
    LPSPropValue p = FindProp(row, PROP_ID(PR_DISPLAY_NAME));
    if (!p) return NULL;
    if (PROP_TYPE(p->ulPropTag) == PT_UNICODE) return p->Value.lpszW;
    if (PROP_TYPE(p->ulPropTag) == PT_STRING8 && p->Value.lpszA) {
        static wchar_t buf[1024];
        if (MultiByteToWideChar(CP_ACP, 0, p->Value.lpszA, -1, buf,
                                (int)(sizeof(buf) / sizeof(buf[0]))) > 0) {
            return buf;
        }
    }
    return NULL;
}

static LPENTRYID GetEntryId(LPSRow row, ULONG* pcb) {
    LPSPropValue p = FindProp(row, PROP_ID(PR_ENTRYID));
    if (!p || PROP_TYPE(p->ulPropTag) != PT_BINARY) return NULL;
    *pcb = p->Value.bin.cb;
    return (LPENTRYID)p->Value.bin.lpb;
}

static bool EntryIdsEqual(ULONG cb1, LPENTRYID eid1, ULONG cb2, LPENTRYID eid2) {
    return cb1 == cb2 && eid1 && eid2 && memcmp(eid1, eid2, cb1) == 0;
}

static void PrintEntryId(const char* label, ULONG cb, LPENTRYID eid) {
    printf("%s: len=%lu hex=", label, (unsigned long)cb);
    const unsigned char* p = (const unsigned char*)eid;
    for (ULONG i = 0; i < cb; ++i) printf("%02x", p[i]);
    printf("\n");
}

static void Usage(void) {
    wprintf(L"Usage:\n");
    wprintf(L"  SetLdapABDefault.exe -list [-p profile]\n");
    wprintf(L"  SetLdapABDefault.exe -set <display name> [-p profile]\n");
    wprintf(L"  SetLdapABDefault.exe -set-index <1-based index> [-p profile]\n");
    wprintf(L"\nIf -p is omitted, the default MAPI profile is used.\n");
    wprintf(L"Exit codes: 0=ok, 1=failure, 2=usage, 3=target not found.\n");
}

int wmain(int argc, wchar_t** argv) {
    if (argc < 2) { Usage(); return RC_USAGE; }

    bool list = false;
    bool set = false;
    bool setIndex = false;
    const wchar_t* target = NULL;
    int targetIndex = 0;
    const wchar_t* profile = NULL;

    for (int i = 1; i < argc; ++i) {
        if (wcscmp(argv[i], L"-list") == 0) {
            list = true;
        } else if (wcscmp(argv[i], L"-set") == 0 && i + 1 < argc) {
            set = true;
            target = argv[++i];
        } else if (wcscmp(argv[i], L"-set-index") == 0 && i + 1 < argc) {
            setIndex = true;
            targetIndex = _wtoi(argv[++i]);
        } else if (wcscmp(argv[i], L"-p") == 0 && i + 1 < argc) {
            profile = argv[++i];
        } else if (wcscmp(argv[i], L"-?") == 0 || wcscmp(argv[i], L"/?") == 0) {
            Usage();
            return RC_OK;
        } else {
            wprintf(L"Unknown argument: %ls\n", argv[i]);
            Usage();
            return RC_USAGE;
        }
    }

    if (!list && !set && !setIndex) { Usage(); return RC_USAGE; }

    MAPIINIT_0 init = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
    HRESULT hr = MAPIInitialize(&init);
    if (FAILED(hr)) { PrintHR("MAPIInitialize", hr); return RC_FAIL; }

    LPMAPISESSION session = NULL;
    ULONG logonFlags = MAPI_EXTENDED | MAPI_NO_MAIL;
    if (!profile) logonFlags |= MAPI_USE_DEFAULT;
    hr = MAPILogonEx(0, (LPTSTR)profile, NULL, logonFlags, &session);
    if (FAILED(hr)) {
        PrintHR("MAPILogonEx", hr);
        MAPIUninitialize();
        return RC_FAIL;
    }

    LPADRBOOK ab = NULL;
    hr = session->OpenAddressBook(0, NULL, 0, &ab);
    if (FAILED(hr)) {
        PrintHR("OpenAddressBook", hr);
        session->Logoff(0, 0, 0);
        session->Release();
        MAPIUninitialize();
        return RC_FAIL;
    }

    LPSRowSet rows = NULL;
    hr = ab->GetSearchPath(0, &rows);
    if (FAILED(hr) || !rows) {
        PrintHR("GetSearchPath", hr);
        ab->Release();
        session->Logoff(0, 0, 0);
        session->Release();
        MAPIUninitialize();
        return RC_FAIL;
    }

    if (list) {
        wprintf(L"Search path entries: %lu\n", (unsigned long)rows->cRows);
        for (ULONG i = 0; i < rows->cRows; ++i) {
            const wchar_t* name = GetDisplayName(&rows->aRow[i]);
            ULONG cb = 0;
            LPENTRYID eid = GetEntryId(&rows->aRow[i], &cb);
            wprintf(L"  [%lu] name=%ls", (unsigned long)(i + 1), name ? name : L"<null>");
            if (eid) {
                printf(" len=%lu hex=", (unsigned long)cb);
                for (ULONG j = 0; j < cb; ++j) printf("%02x", ((unsigned char*)eid)[j]);
                printf("\n");
            } else {
                printf(" entryid=<missing>\n");
            }
        }
        ULONG cbDef = 0;
        LPENTRYID def = NULL;
        hr = ab->GetDefaultDir(&cbDef, &def);
        if (SUCCEEDED(hr) && def) {
            wprintf(L"DefaultDir:\n");
            PrintEntryId("  eid", cbDef, def);
            MAPIFreeBuffer(def);
        } else {
            wprintf(L"DefaultDir: <none> hr=0x%08lx\n", (unsigned long)hr);
        }
    }

    int exitCode = RC_OK;
    if (set || setIndex) {
        ULONG found = (ULONG)-1;
        for (ULONG i = 0; i < rows->cRows; ++i) {
            const wchar_t* name = GetDisplayName(&rows->aRow[i]);
            if (set && name && wcscmp(name, target) == 0) {
                found = i;
                break;
            }
            if (setIndex && (int)(i + 1) == targetIndex) {
                found = i;
                break;
            }
        }
        if (found == (ULONG)-1) {
            wprintf(L"Target address list not found in search path.\n");
            exitCode = RC_NOT_FOUND;
        } else {
            ULONG cb = 0;
            LPENTRYID eid = GetEntryId(&rows->aRow[found], &cb);
            if (!eid) {
                wprintf(L"Target row has no PR_ENTRYID.\n");
                exitCode = RC_FAIL;
            } else {
                const wchar_t* foundName = GetDisplayName(&rows->aRow[found]);
                wprintf(L"Target: %ls\n", foundName ? foundName : L"<unnamed>");
                PrintEntryId("Target entryid", cb, eid);

                // Check the current default and search order.
                bool defaultMatches = false;
                ULONG cbDef = 0;
                LPENTRYID def = NULL;
                hr = ab->GetDefaultDir(&cbDef, &def);
                if (SUCCEEDED(hr) && def) {
                    defaultMatches = EntryIdsEqual(cb, eid, cbDef, def);
                    MAPIFreeBuffer(def);
                }

                bool firstMatches = false;
                if (rows->cRows > 0) {
                    ULONG cbFirst = 0;
                    LPENTRYID first = GetEntryId(&rows->aRow[0], &cbFirst);
                    firstMatches = EntryIdsEqual(cb, eid, cbFirst, first);
                }

                if (defaultMatches && firstMatches) {
                    wprintf(L"Already default and first in search path.\n");
                } else {
                    if (!defaultMatches) {
                        hr = ab->SetDefaultDir(cb, eid);
                        wprintf(L"SetDefaultDir: hr=0x%08lx\n", (unsigned long)hr);
                        if (FAILED(hr)) exitCode = RC_FAIL;
                    } else {
                        wprintf(L"DefaultDir already matches target.\n");
                    }

                    if (!firstMatches) {
                        LPSRowSet newRows = NULL;
                        SIZE_T bytes = sizeof(SRowSet) + (rows->cRows > 1 ? (rows->cRows - 1) * sizeof(SRow) : 0);
                        hr = MAPIAllocateBuffer((ULONG)bytes, (LPVOID*)&newRows);
                        if (SUCCEEDED(hr)) {
                            ZeroMemory(newRows, bytes);
                            newRows->cRows = rows->cRows;
                            bool ok = true;
                            for (ULONG i = 0; i < rows->cRows && ok; ++i) {
                                ULONG src = (i == 0) ? found : (i <= found ? i - 1 : i);
                                LPSRow oldRow = &rows->aRow[src];
                                newRows->aRow[i].cValues = oldRow->cValues;
                                SIZE_T propBytes = oldRow->cValues * sizeof(SPropValue);
                                hr = MAPIAllocateBuffer((ULONG)propBytes, (LPVOID*)&newRows->aRow[i].lpProps);
                                if (FAILED(hr)) { ok = false; break; }
                                CopyMemory(newRows->aRow[i].lpProps, oldRow->lpProps, propBytes);
                                for (ULONG j = 0; j < oldRow->cValues; ++j) {
                                    LPSPropValue srcProp = &oldRow->lpProps[j];
                                    LPSPropValue dstProp = &newRows->aRow[i].lpProps[j];
                                    if (PROP_TYPE(srcProp->ulPropTag) == PT_BINARY && srcProp->Value.bin.lpb) {
                                        void* p = NULL;
                                        if (SUCCEEDED(MAPIAllocateBuffer(srcProp->Value.bin.cb, &p))) {
                                            CopyMemory(p, srcProp->Value.bin.lpb, srcProp->Value.bin.cb);
                                            dstProp->Value.bin.lpb = (LPBYTE)p;
                                        } else {
                                            ok = false;
                                        }
                                    } else if (PROP_TYPE(srcProp->ulPropTag) == PT_UNICODE && srcProp->Value.lpszW) {
                                        SIZE_T size = (wcslen(srcProp->Value.lpszW) + 1) * sizeof(wchar_t);
                                        void* p = NULL;
                                        if (SUCCEEDED(MAPIAllocateBuffer((ULONG)size, &p))) {
                                            CopyMemory(p, srcProp->Value.lpszW, size);
                                            dstProp->Value.lpszW = (LPWSTR)p;
                                        } else {
                                            ok = false;
                                        }
                                    } else if (PROP_TYPE(srcProp->ulPropTag) == PT_STRING8 && srcProp->Value.lpszA) {
                                        SIZE_T size = strlen(srcProp->Value.lpszA) + 1;
                                        void* p = NULL;
                                        if (SUCCEEDED(MAPIAllocateBuffer((ULONG)size, &p))) {
                                            CopyMemory(p, srcProp->Value.lpszA, size);
                                            dstProp->Value.lpszA = (LPSTR)p;
                                        } else {
                                            ok = false;
                                        }
                                    }
                                }
                            }
                            if (ok) {
                                hr = ab->SetSearchPath(0, newRows);
                                wprintf(L"SetSearchPath: hr=0x%08lx\n", (unsigned long)hr);
                                if (FAILED(hr)) exitCode = RC_FAIL;
                            } else {
                                wprintf(L"SetSearchPath buffer copy failed.\n");
                                exitCode = RC_FAIL;
                            }
                            for (ULONG i = 0; i < newRows->cRows; ++i) {
                                if (newRows->aRow[i].lpProps) {
                                    for (ULONG j = 0; j < newRows->aRow[i].cValues; ++j) {
                                        LPSPropValue p = &newRows->aRow[i].lpProps[j];
                                        if (PROP_TYPE(p->ulPropTag) == PT_BINARY && p->Value.bin.lpb) {
                                            MAPIFreeBuffer(p->Value.bin.lpb);
                                        } else if (PROP_TYPE(p->ulPropTag) == PT_UNICODE && p->Value.lpszW) {
                                            MAPIFreeBuffer(p->Value.lpszW);
                                        } else if (PROP_TYPE(p->ulPropTag) == PT_STRING8 && p->Value.lpszA) {
                                            MAPIFreeBuffer(p->Value.lpszA);
                                        }
                                    }
                                    MAPIFreeBuffer(newRows->aRow[i].lpProps);
                                }
                            }
                            MAPIFreeBuffer(newRows);
                        } else {
                            wprintf(L"MAPIAllocateBuffer for search path failed: hr=0x%08lx\n", (unsigned long)hr);
                            exitCode = RC_FAIL;
                        }
                    } else {
                        wprintf(L"Search path already has target first.\n");
                    }
                }

                // Re-read and verify both settings.
                if (exitCode == RC_OK) {
                    ULONG cbDef2 = 0;
                    LPENTRYID def2 = NULL;
                    bool verifiedDefault = false;
                    hr = ab->GetDefaultDir(&cbDef2, &def2);
                    if (SUCCEEDED(hr) && def2) {
                        verifiedDefault = EntryIdsEqual(cb, eid, cbDef2, def2);
                        MAPIFreeBuffer(def2);
                    }
                    LPSRowSet rows2 = NULL;
                    bool verifiedFirst = false;
                    hr = ab->GetSearchPath(0, &rows2);
                    if (SUCCEEDED(hr) && rows2 && rows2->cRows > 0) {
                        ULONG cbFirst2 = 0;
                        LPENTRYID first2 = GetEntryId(&rows2->aRow[0], &cbFirst2);
                        verifiedFirst = EntryIdsEqual(cb, eid, cbFirst2, first2);
                    }
                    FreeRowSetManually(rows2);
                    if (verifiedDefault && verifiedFirst) {
                        wprintf(L"VERIFIED default and search order.\n");
                    } else {
                        wprintf(L"VERIFY FAILED default=%d first=%d\n", verifiedDefault ? 1 : 0, verifiedFirst ? 1 : 0);
                        exitCode = RC_FAIL;
                    }
                }
            }
        }
    }

    FreeRowSetManually(rows);
    ab->Release();
    session->Logoff(0, 0, 0);
    session->Release();
    MAPIUninitialize();
    return exitCode;
}
