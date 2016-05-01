
// viewDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "view.h"
#include "viewDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CviewDlg 对话框




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
END_MESSAGE_MAP()


// CviewDlg 消息处理程序

BOOL CviewDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	m_exp.put_MenuBar(FALSE);
	m_exp.SetSilent(TRUE);
	m_exp.Navigate(_T("C:\\tmp\\1.txt"), NULL, NULL, NULL, NULL);
	CRect rectClient;
	GetClientRect(&rectClient);
	m_exp.MoveWindow(&rectClient);
	//m_exp.Navigate(_T("C:\\u\\hh.mht"), NULL, NULL, NULL, NULL);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CviewDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CviewDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

