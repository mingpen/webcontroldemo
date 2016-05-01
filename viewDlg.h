
// viewDlg.h : ͷ�ļ�
//

#pragma once
#include "explorer1.h"


// CviewDlg �Ի���
class CviewDlg : public CDialog
{
// ����
public:
	CviewDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_VIEW_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	CString m_strFilePathView;

	CString m_strFilePathPlain;

	// ���ɵ���Ϣӳ�亯��
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
