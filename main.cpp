#define UNICODE
#pragma comment(lib,"shlwapi")
#pragma comment(lib,"iphlpapi")
#pragma comment(lib,"ws2_32")
#include<winsock2.h>
#include<windows.h>
#include<stdio.h>
#include<shlwapi.h>
#include<shlobj.h>
#include<iphlpapi.h>
#include"resource.h"

#define MACLEN (6)
#define PACKLEN (17*MACLEN)
#define WINDOW_WIDTH (400)
#define WINDOW_HEIGHT (300)
enum{
	ID_WAKE=100,
	ID_NEW=101,
	ID_EDIT=102,
	ID_DELETE=103
};

TCHAR szClassName[]=TEXT("Window");

class DATA
{
public:
	BOOL bSetMac;
	TCHAR szName[256];
	TCHAR szMacAddress[18];
	BYTE byMac[MACLEN];
	DATA():bSetMac(0)
	{
		szName[0]=0;
		szMacAddress[0]=0;
		ZeroMemory(byMac,MACLEN);
	}
	void SetName(LPCTSTR lpString)
	{
		lstrcpy(szName,lpString);
	}
	LPCTSTR GetName()
	{
		return szName;
	}
	LPCTSTR GetMac()
	{
		return szMacAddress;
	}
	BOOL SetMac(LPCTSTR lpString)
	{
		TCHAR szString[13];
		int i=0;
		int j=0;
		while(1)
		{
			if(lpString[i]==0)break;
			if(iswxdigit(lpString[i]))
			{
				if(j>=12)return FALSE;
				szString[j]=towupper(lpString[i]);
				j++;
			}
			i++;
		}
		szString[j]=0;
		if(lstrlen(szString)!=12)return FALSE;
		for(i=0;i<MACLEN;i++)
		{
			const TCHAR x[]={szString[i*2],szString[i*2+1],0};
			byMac[i]=(BYTE)wcstoul(x,0,16);
		}
		wsprintf(szMacAddress,TEXT("%02X-%02X-%02X-%02X-%02X-%02X"),byMac[0],byMac[1],byMac[2],byMac[3],byMac[4],byMac[5]);
		bSetMac=TRUE;
		return TRUE;
	}
	void WakeOnLan()
	{
		if(!bSetMac)return;
		struct sockaddr_in addr;
		int t=0,i;
		SOCKET sock;
		BOOL BroadCast=1;
		CHAR pack[PACKLEN];
		WSADATA wsaData;
		WSAStartup(MAKEWORD(2,2),&wsaData);
		while(t<MACLEN)
		{
			pack[t]=(CHAR)0xFF;
			t++;
		}
		for(i=0;i<MACLEN;i++)
		{
			pack[t]=byMac[i];
			t++;
		}
		for(;t<PACKLEN;++t)
		{
			pack[t]=pack[t-MACLEN];
		}
		sock=socket(AF_INET,SOCK_DGRAM,IPPROTO_UDP);
		if(sock==INVALID_SOCKET)
		{
			WSAGetLastError();
		}
		else
		{
			setsockopt(sock,SOL_SOCKET,SO_BROADCAST,(const char*)&BroadCast,sizeof(BroadCast));
			addr.sin_family=AF_INET;
			addr.sin_addr.s_addr=0xFFFFFFFF;
			addr.sin_port=htons(9);
			t=sendto(sock,pack,PACKLEN,0,(struct sockaddr*)&addr,sizeof(addr));
			if(t!=PACKLEN)
			{
				WSAGetLastError();
			}
			closesocket(sock);
		}
		WSACleanup();
	}
	static LPSTR WidetoShiftJIS(LPCWSTR lpszString)
	{
        const int cchMultiByte=WideCharToMultiByte(CP_ACP,0,lpszString,-1,0,0,0,0);
        LPSTR lpa=new CHAR[cchMultiByte];
        if(lpa==0)return 0;
        *lpa='\0';
        const int nMultiCount=WideCharToMultiByte(CP_ACP,0,lpszString,-1,lpa,cchMultiByte,0,0);
        if(nMultiCount<=0)
		{
            delete[]lpa;
            return 0;
        }
		return lpa;
 	}
	static LPWSTR ShiftJIStoWide(LPSTR lpszString)
	{
        const int cchMultiByte=MultiByteToWideChar(CP_ACP,0,lpszString,-1,0,0);
        LPWSTR lpa=new WCHAR[cchMultiByte];
        if(lpa==0)return 0;
        *lpa='\0';
        const int nMultiCount=MultiByteToWideChar(CP_ACP,0,lpszString,-1,lpa,cchMultiByte);
        if(nMultiCount<=0)
		{
            delete[]lpa;
            return 0;
        }
		return lpa;
 	}
	static void Read(HWND hList,LPCTSTR lpszFileName)
	{
		CHAR *token1;
		CHAR *token2;
		const CHAR seps[]=",";
		CHAR buf[1024];
		FILE *fp;
		if(_wfopen_s(&fp,lpszFileName,L"r")!=0)
		{
			return;
		}
		while(fgets(buf,1024,fp))
		{
			CHAR *ctx;
			token1=strtok_s(buf,seps,&ctx);
			token2=strtok_s(0,seps,&ctx);
			if(token1&&token2)
			{
				DATA*pData=new DATA;
				LPCWSTR p=DATA::ShiftJIStoWide(token1);
				if(p)
				{
					pData->SetName(p);
					delete[]p;
				}
				p=DATA::ShiftJIStoWide(token2);
				if(p)
				{
					pData->SetMac(p);
					delete[]p;
				}
				const int nIndex=SendMessage(hList,LB_ADDSTRING,0,(long)pData->GetName());
				SendMessage(hList,LB_SETITEMDATA,nIndex,(long)pData);
			}
		}
		fclose(fp);
	}
	static void Save(HWND hList,LPCTSTR lpszFileName)
	{
		DWORD dwAccBytes;
		LPSTR p;
		const HANDLE hFile=CreateFile(lpszFileName,
			GENERIC_WRITE,
			0,
			0,
			CREATE_ALWAYS,
			FILE_ATTRIBUTE_NORMAL,
			0);
		const int nCount=SendMessage(hList,LB_GETCOUNT,0,0);
		for(int i=0;i<nCount;i++)
		{
			DATA* pData=(DATA*)SendMessage(hList,LB_GETITEMDATA,i,0);
			if(pData)
			{
				p=WidetoShiftJIS(pData->GetName());
				if(p)
				{
					WriteFile(hFile,
						p,
						lstrlenA(p),
						&dwAccBytes,
						0);
					delete[]p;
				}
				WriteFile(hFile,
					",",
					1,
					&dwAccBytes,
					0);
				p=WidetoShiftJIS(pData->GetMac());
				if(p)
				{
					WriteFile(hFile,
						p,
						lstrlenA(p),
						&dwAccBytes,
						0);
					delete[]p;
				}
				WriteFile(hFile,
					"\r\n",
					2,
					&dwAccBytes,
					0);
				
				SendMessage(hList,LB_SETITEMDATA,i,0);
				delete pData;
			}
		}
		CloseHandle(hFile);
	}
};


void ShowNetwork(HWND hWnd)
{
	LPMALLOC pMalloc;
	if(SHGetMalloc(&pMalloc)!=NOERROR)
	{
		return;
	}
	TCHAR szPath[MAX_PATH];
	ZeroMemory(szPath,sizeof(TCHAR)*MAX_PATH);
	LPITEMIDLIST pidlRoot,pidlBrowse;
	SHGetSpecialFolderLocation(hWnd,CSIDL_NETWORK,&pidlRoot);
	BROWSEINFO bi;
	ZeroMemory(&bi,sizeof(BROWSEINFO));
	bi.hwndOwner=hWnd;
	bi.pidlRoot=pidlRoot;
	bi.pszDisplayName=szPath;
	bi.lpszTitle=TEXT("コンピューターを選択してください。");
	bi.ulFlags=BIF_BROWSEFORCOMPUTER;
	pidlBrowse=SHBrowseForFolder(&bi);
	if(pidlBrowse)
	{
		SetDlgItemText(hWnd,IDC_EDIT1,szPath);
		LPSTR hostname=DATA::WidetoShiftJIS(szPath);
		if(hostname)
		{
			struct hostent *host;
			if((host=gethostbyname(hostname))!=0)
			{
				IPAddr ip;
				memcpy(&ip,host->h_addr,sizeof(ip));
				unsigned char mac[MACLEN]={'\0'};
				ULONG ulLen=sizeof(mac)/sizeof(mac[0]);
				DWORD dwRet;
				dwRet=SendARP(ip,0,(PULONG)mac,&ulLen);
				if(dwRet==NO_ERROR)
				{
					TCHAR buf[18];
					wsprintf(buf,
						TEXT("%02X-%02X-%02X-%02X-%02X-%02X"),
						mac[0],mac[1],mac[2],
						mac[3],mac[4],mac[5]);
					SetDlgItemText(hWnd,IDC_EDIT2,buf);
				} 
			}
			delete[]hostname;
		}
	}
	if(pidlRoot!=0)
		pMalloc->Free(pidlRoot);
	if(pidlBrowse!=0)
		pMalloc->Free(pidlBrowse);
	pMalloc->Release();
}


int CALLBACK DialogProc(HWND hDlg,unsigned msg,WPARAM wParam,LPARAM lParam)
{
	static DATA *pData;
	switch(msg)
	{
	case WM_INITDIALOG:
		pData=(DATA*)lParam;
		if(pData)
		{
			SetDlgItemText(hDlg,IDC_EDIT1,pData->GetName());
			SetDlgItemText(hDlg,IDC_EDIT2,pData->GetMac());
		}
		return TRUE;
	case WM_COMMAND:
		switch(LOWORD(wParam))
		{
		case IDOK:
			{
				const HWND hEdit1=GetDlgItem(hDlg,IDC_EDIT1);
				const HWND hEdit2=GetDlgItem(hDlg,IDC_EDIT2);
				const int nCount=GetWindowTextLength(hEdit1);
				if(!nCount)
				{
					SetFocus(hEdit1);
					return FALSE;
				}
				if(!GetWindowTextLength(hEdit2))
				{
					SetFocus(hEdit2);
					return FALSE;
				}
				TCHAR szName[256];
				TCHAR szMac[256];
				GetDlgItemText(hDlg,IDC_EDIT1,szName,256);
				for(int i=0;i<nCount;i++)
				{
					if(szName[i]==TEXT(','))
					{
						MessageBox(hDlg,TEXT("名前にコンマ(,)は使用できません。"),0,MB_ICONERROR);
						SetFocus(hEdit1);
						return FALSE;
					}
				}
				GetDlgItemText(hDlg,IDC_EDIT2,szMac,256);
				if(!pData->SetMac(szMac))
				{
					SendMessage(hEdit2,EM_SETSEL,0,-1);
					SetFocus(hEdit2);
					return FALSE;
				}
				pData->SetName(szName);
				EndDialog(hDlg,LOWORD(wParam));
			}
			break;
		case IDCANCEL:
			EndDialog(hDlg,LOWORD(wParam));
			break;
		case IDC_REFER:
			ShowNetwork(hDlg);
			break;
		}
		break;
	}
	return FALSE;
}


LRESULT CALLBACK WndProc(HWND hWnd,UINT msg,WPARAM wParam,LPARAM lParam)
{
	static HFONT hFont;
	static HWND hList;
	static HWND hButton1;
	static HWND hButton2;
	static HWND hButton3;
	static HWND hButton4;
	static HWND hStatic;
	static TCHAR szFile[MAX_PATH];
	static HICON hIcon;
	switch(msg)
	{
	case WM_CREATE:
		hFont=CreateFont(-14,0,0,0,0,0,0,0,0,0,0,0,0,TEXT("メイリオ"));
		hIcon=(HICON)LoadImage(((LPCREATESTRUCT)lParam)->hInstance,MAKEINTRESOURCE(IDI_ICON1),IMAGE_ICON,48,48,0);
		hList=CreateWindowEx(WS_EX_CLIENTEDGE,TEXT("LISTBOX"),0,WS_VISIBLE|WS_CHILD|WS_VSCROLL|WS_TABSTOP|LBS_NOINTEGRALHEIGHT|LBS_OWNERDRAWFIXED|LBS_NOTIFY,10,10,256,256,hWnd,(HMENU)99,((LPCREATESTRUCT)lParam)->hInstance,0);
		SendMessage(hList,LB_SETITEMHEIGHT,0,MAKELPARAM(64,0));
		GetModuleFileName(((LPCREATESTRUCT)lParam)->hInstance,szFile,MAX_PATH);
		PathStripPath(szFile);
		szFile[lstrlen(szFile)-4]=0;
		lstrcat(szFile,TEXT(".csv"));
		DATA::Read(hList,szFile);
		hButton1=CreateWindow(TEXT("BUTTON"),TEXT("起動(&W)"),WS_VISIBLE|WS_CHILD|WS_TABSTOP|WS_DISABLED,0,0,0,0,hWnd,(HMENU)ID_WAKE,((LPCREATESTRUCT)lParam)->hInstance,0);
		hButton2=CreateWindow(TEXT("BUTTON"),TEXT("新規(&N)..."),WS_VISIBLE|WS_CHILD|WS_TABSTOP,0,0,0,0,hWnd,(HMENU)ID_NEW,((LPCREATESTRUCT)lParam)->hInstance,0);
		hButton3=CreateWindow(TEXT("BUTTON"),TEXT("編集(&E)..."),WS_VISIBLE|WS_CHILD|WS_TABSTOP|WS_DISABLED,0,0,0,0,hWnd,(HMENU)ID_EDIT,((LPCREATESTRUCT)lParam)->hInstance,0);
		hButton4=CreateWindow(TEXT("BUTTON"),TEXT("削除(&D)"),WS_VISIBLE|WS_CHILD|WS_TABSTOP|WS_DISABLED,0,0,0,0,hWnd,(HMENU)ID_DELETE,((LPCREATESTRUCT)lParam)->hInstance,0);
		hStatic=CreateWindow(TEXT("STATIC"),TEXT("Ver : ")
			TEXT(__DATE__),WS_VISIBLE|WS_CHILD|WS_DISABLED|SS_RIGHT,0,0,0,0,hWnd,0,((LPCREATESTRUCT)lParam)->hInstance,0);
		SendMessage(hList,WM_SETFONT,(WPARAM)hFont,0);
		SendMessage(hButton1,WM_SETFONT,(WPARAM)hFont,0);
		SendMessage(hButton2,WM_SETFONT,(WPARAM)hFont,0);
		SendMessage(hButton3,WM_SETFONT,(WPARAM)hFont,0);
		SendMessage(hButton4,WM_SETFONT,(WPARAM)hFont,0);
		SendMessage(hStatic,WM_SETFONT,(WPARAM)hFont,0);
		break;
	case WM_SIZE:
		MoveWindow(hList,10,10,LOWORD(lParam)-158,HIWORD(lParam)-20,TRUE);
		MoveWindow(hButton1,LOWORD(lParam)-138,10,128,30,TRUE);
		MoveWindow(hButton2,LOWORD(lParam)-138,50,128,30,TRUE);
		MoveWindow(hButton3,LOWORD(lParam)-138,90,128,30,TRUE);
		MoveWindow(hButton4,LOWORD(lParam)-138,130,128,30,TRUE);
		MoveWindow(hStatic,LOWORD(lParam)-138,HIWORD(lParam)-30,128,20,TRUE);
		break;
	case WM_DRAWITEM:
		if((UINT)wParam==99)
		{
			LPDRAWITEMSTRUCT lpdis=(LPDRAWITEMSTRUCT)lParam;
			DATA* pData=(DATA*)lpdis->itemData;
			if(pData)
			{
				HBRUSH hBrush;
				COLORREF colText;
				COLORREF colSubText;
				if(lpdis->itemState&ODS_SELECTED)
				{
					hBrush=CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT));
					colText=GetSysColor(COLOR_HIGHLIGHTTEXT);
					colSubText=(GetSysColor(COLOR_HIGHLIGHTTEXT)+GetSysColor(COLOR_HIGHLIGHT))/2;
				}
				else
				{
					hBrush=CreateSolidBrush(GetSysColor(COLOR_WINDOW));
					colText=GetSysColor(COLOR_WINDOWTEXT);
					colSubText=GetSysColor(COLOR_BTNSHADOW);
				}
				FillRect(lpdis->hDC,&lpdis->rcItem,hBrush);
				DrawIconEx(lpdis->hDC,lpdis->rcItem.left+4,lpdis->rcItem.top+8,hIcon,48,48,0,0,DI_IMAGE|DI_MASK);
				if(lpdis->itemState&ODS_FOCUS)
				{
					DrawFocusRect(lpdis->hDC,&lpdis->rcItem);
				}
				DeleteObject(hBrush);
				SetBkMode(lpdis->hDC,TRANSPARENT);
				LPCTSTR lpName=pData->GetName();
				LPCTSTR lpMac=pData->GetMac();
				SetTextColor(lpdis->hDC,colText);
				TextOut(lpdis->hDC,lpdis->rcItem.left+58,lpdis->rcItem.top+7,lpName,lstrlen(lpName));
				SetTextColor(lpdis->hDC,colSubText);
				TextOut(lpdis->hDC,lpdis->rcItem.left+58,lpdis->rcItem.top+33,lpMac,lstrlen(lpMac));
			}
		}
		break;
	case WM_COMMAND:
		switch(LOWORD(wParam))
		{
		case IDCANCEL:
			SendMessage(hWnd,WM_CLOSE,0,0);
			break;
		case 99:
			switch(HIWORD(wParam))
			{
			case LBN_DBLCLK:
				SendMessage(hWnd,WM_COMMAND,ID_WAKE,0);
				break;
			case LBN_SELCHANGE:
				{
					const BOOL bEnable=(SendMessage(hList,LB_GETCURSEL,0,0)!=LB_ERR);
					EnableWindow(hButton1,bEnable);
					EnableWindow(hButton3,bEnable);
					EnableWindow(hButton4,bEnable);
				}
				break;
			}
			break;
		case ID_WAKE:
			{
				const int nIndex=SendMessage(hList,LB_GETCURSEL,0,0);
				if(nIndex!=LB_ERR)
				{
					DATA* pData=(DATA*)SendMessage(hList,LB_GETITEMDATA,nIndex,0);
					if(pData)
					{
						pData->WakeOnLan();
					}
				}
			}
			break;
		case ID_NEW:
			{
				DATA* pData=new DATA;
				if(DialogBoxParam(GetModuleHandle(0),MAKEINTRESOURCE(IDD_DIALOG),hWnd,DialogProc,(long)pData)==IDOK)
				{
					const int nIndex=SendMessage(hList,LB_ADDSTRING,0,(long)pData->GetName());
					SendMessage(hList,LB_SETITEMDATA,nIndex,(long)pData);
					SendMessage(hList,LB_SETCURSEL,nIndex,0);
					PostMessage(hWnd,WM_COMMAND,MAKEWPARAM(99,LBN_SELCHANGE),0);
					SetFocus(hList);
				}
				else
				{
					delete pData;
				}
			}
			break;
		case ID_EDIT:
			{
				const int nIndex=SendMessage(hList,LB_GETCURSEL,0,0);
				if(nIndex!=LB_ERR)
				{
					DATA* pData=(DATA*)SendMessage(hList,LB_GETITEMDATA,nIndex,0);
					if(pData)
					{
						if(DialogBoxParam(GetModuleHandle(0),MAKEINTRESOURCE(IDD_DIALOG),hWnd,DialogProc,(long)pData)==IDOK)
						{
							InvalidateRect(hList,0,1);
						}
					}
				}
			}
			break;
		case ID_DELETE:
			{
				const int nIndex=SendMessage(hList,LB_GETCURSEL,0,0);
				if(nIndex!=LB_ERR)
				{
					DATA* pData=(DATA*)SendMessage(hList,LB_GETITEMDATA,nIndex,0);
					if(pData)
					{
						delete pData;
						SendMessage(hList,LB_DELETESTRING,nIndex,0);
						PostMessage(hWnd,WM_COMMAND,MAKEWPARAM(99,LBN_SELCHANGE),0);
					}
				}
			}
			break;
		}
		break;
	case WM_GETMINMAXINFO:
		{
			MINMAXINFO* lpMMI=(MINMAXINFO*)lParam;
			lpMMI->ptMinTrackSize.x=WINDOW_WIDTH;
			lpMMI->ptMinTrackSize.y=WINDOW_HEIGHT;
		}
		break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		DATA::Save(hList,szFile);
		DeleteObject(hFont);
		PostQuitMessage(0);
		break;
	default:
		return(DefDlgProc(hWnd,msg,wParam,lParam));
	}
	return 0;
}


int WINAPI WinMain(HINSTANCE hInstance,HINSTANCE hPreInst,
				   LPSTR pCmdLine,int nCmdShow)
{
	MSG msg;
	const WNDCLASS wndclass={0,WndProc,0,DLGWINDOWEXTRA,hInstance,0,LoadCursor(0,IDC_ARROW),0,0,szClassName};
	RegisterClass(&wndclass);
	const HWND hWnd=CreateWindow(szClassName,
		TEXT("Wake On Lan"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		WINDOW_WIDTH,
		WINDOW_HEIGHT,
		0,
		0,
		hInstance,
		0);
	ShowWindow(hWnd,nCmdShow);
	UpdateWindow(hWnd);
	ACCEL Accel[]={{FVIRTKEY,VK_DELETE,ID_DELETE}};
	HACCEL hAccel=CreateAcceleratorTable(Accel,1);
	while(GetMessage(&msg,0,0,0))
	{
		if(!TranslateAccelerator(hWnd,hAccel,&msg)&&!IsDialogMessage(hWnd,&msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	DestroyAcceleratorTable(hAccel);
	return msg.wParam;
}
