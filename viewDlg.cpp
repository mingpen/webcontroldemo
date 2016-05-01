
// viewDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "msword.h"

#include "view.h"
#include "viewDlg.h"
#include "file/FileHelper.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CviewDlg �Ի���




CviewDlg::CviewDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CviewDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CviewDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EXPLORER1, m_exp);
}

BEGIN_MESSAGE_MAP(CviewDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_WM_TIMER()
	ON_WM_DESTROY()
END_MESSAGE_MAP()


// CviewDlg ��Ϣ�������

BOOL CviewDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	ConvertToMHT();
	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	m_exp.put_MenuBar(FALSE);
	m_exp.SetSilent(TRUE);
	m_exp.Navigate(m_strFilePathView, NULL, NULL, NULL, NULL);
	CRect rectClient;
	GetClientRect(&rectClient);
	m_exp.MoveWindow(&rectClient);
	m_nTimer = SetTimer(1, 1000, NULL);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CviewDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CviewDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

bool CviewDlg::ConvertToMHT()
{
	//
	m_strFilePathPlain = _T("D:\\t\\TEST.docx");
	m_strFilePathView = _T("C:\\1\\1.mht");
	//
	bool bRet = false;

	_Application app;//����һ��wordӦ�ö���
	do 
	{
		if(!app.CreateDispatch(_T("Word.Application")))  //����word
		{
			MessageBox(_T("û�ɹ���װMS WORD,�޷��鿴word�ĵ���"));
			return bRet;
		}
		app.SetVisible(FALSE);

		Documents docs=app.GetDocuments();
		CComVariant FileName(m_strFilePathPlain.GetString());
		CComVariant ConfirmConversions(false),ReadOnly(TRUE);
		CComVariant AddToRecentFiles(false),PasswordDocument(_T(""));
		CComVariant PasswordTemplate(_T("")),Revert(false),WritePasswordDocument(_T(""));
		CComVariant WritePasswordTemplate(_T("")),Format(0),Encoding(false);
		CComVariant Visible(true),OpenAndRepair(false),DocumentDirection(false);
		CComVariant NoEncodingDialog(false),XMLTransform(_T(""));

		_Document doc=docs.Open(&FileName, &ConfirmConversions, &ReadOnly, 
			&AddToRecentFiles, &PasswordDocument, 
			&PasswordTemplate, &Revert, 
			&WritePasswordDocument, &WritePasswordTemplate, 
			&Format, &Encoding, &Visible, 
			&OpenAndRepair, &DocumentDirection, 
			&NoEncodingDialog, &XMLTransform);
		if ( NULL  != doc )
		{
			doc.SetSaveEncoding(3);
			COleVariant vSaveFormat((short)9),vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
			doc.SaveAs(COleVariant(m_strFilePathView.GetString()), vSaveFormat, vOpt, vOpt, vOpt,
				vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt);		
		}
		else {
			MessageBox(_T("��word�ĵ�ʧ�ܣ�"));
		}

		if ( File_GetSize(m_strFilePathView) > 0 )
		{
			bRet = true;
		}

		docs.ReleaseDispatch();

	} while (FALSE);
	//�˳�
	CComVariant SaveChanges(false),OriginalFormat,RounteDocument;
	app.Quit(&SaveChanges,&OriginalFormat,&RounteDocument);
	app.ReleaseDispatch();

	return bRet;
}



void CviewDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default
	long s = m_exp.get_ReadyState();
	TRACKINFO(L"get_ReadyState=%d", s);
	if ( READYSTATE_COMPLETE == s )
	{
		KillTimer(m_nTimer);
		m_nTimer = 0;
		File_Remove(m_strFilePathView.GetString());
	}


	CDialog::OnTimer(nIDEvent);
}

void CviewDlg::OnDestroy()
{
	CDialog::OnDestroy();

	// TODO: Add your message handler code here
	File_Remove(m_strFilePathView.GetString());

}
