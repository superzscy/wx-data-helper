#pragma once

#include <wx/wx.h>
#include <wx/choice.h>
#include <wx/textctrl.h>
#include <xlnt/xlnt.hpp>

class ExcelInputPanel : public wxPanel
{
public:
    ExcelInputPanel(wxWindow* parent, const wxString& labelPrefix, bool bHasReturnCol);

    // 获取信息接口
    wxString GetFilePath() const;
    wxString GetSelectedSheet() const;

    bool IsReady() const;

    bool Parse();

    wxString GetMatchColumnName() const;
    wxString GetReturnColumnName() const;
    wxString GetErrorMsg() const { return m_errorMsg; }
    wxString GetOutputFilePath() const { return m_outputFilePath; }

    const wxBaseArray<wxArrayString>& GetContentRows() const { return m_contentRows; }
    const int GetMatchColumnIndex() const { return m_matchColumnIndex; }
    const int GetReturnColumnIndex() const { return m_returnColumnIndex; }
    bool Match(const wxBaseArray<wxArrayString>& contentRows, int matchColumnIndex, int returnColumnIndex);

private:
    void OnSelectFile(wxCommandEvent& event);

    wxButton* m_btnSelectFile = nullptr;
    wxTextCtrl* m_txtFilePath = nullptr;
    wxChoice* m_sheetChoice = nullptr;
    wxTextCtrl* m_columnInputMatch = nullptr;
    wxTextCtrl* m_columnInputReturn = nullptr;

    wxString m_labelPrefix;
    bool m_loaded = false;
    xlnt::workbook m_workbook;

	bool m_bHasReturnCol = false;
    int m_matchColumnIndex = -1;
	int m_returnColumnIndex = -1;
    int m_headerRowIndex = -1;
    wxBaseArray<wxArrayString> m_contentRows;
    wxArrayInt m_contentRowNumbers;

    wxString m_errorMsg;
    wxString m_outputFilePath;
};
