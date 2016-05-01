
// viewDlg.h : 头文件
//

#pragma once
#include "explorer1.h"


// CviewDlg 对话框
class CviewDlg : public CDialog
{
// 构造
public:
	CviewDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_VIEW_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	CString m_strFilePathView;

	CString m_strFilePathPlain;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CExplorer1 m_exp;

	bool ConvertToMHT();
	afx_msg void OnTimer(UINT_PTR nIDEvent);

private:
	UINT_PTR m_nTimer;
public:
	afx_msg void OnDestroy();
};
