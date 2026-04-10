#pragma once

#include <wx/wx.h>
#include <wx/choice.h>
#include <wx/textctrl.h>
#include <xlnt/xlnt.hpp>
#include <vector>

class ExcelInputPanel : public wxPanel
{
public:
    ExcelInputPanel(wxWindow* parent, const wxString& labelPrefix, bool bHasReturnCol);

    // Basic accessors
    wxString GetFilePath() const;
    wxString GetSelectedSheet() const;

    bool IsReady() const;

    bool Parse();

    wxString GetMatchColumnName() const;
    wxString GetExtractColumnName() const;
    wxArrayString GetExtractColumnNames() const;
    const wxArrayString& GetTitleNames() const { return m_titleNames; }
    void ConfigureFilterBaseColumns(const wxArrayString& baseColumns);
    bool ShouldFilter() const { return m_filterLeftColumnIndex >= 0 && m_filterRightColumnIndex >= 0; }
    int GetFilterLeftColumnIndex() const { return m_filterLeftColumnIndex; }
    int GetFilterRightColumnIndex() const { return m_filterRightColumnIndex; }
    wxString GetErrorMsg() const { return m_errorMsg; }
    wxString GetOutputFilePath() const { return m_outputFilePath; }

    const wxBaseArray<wxArrayString>& GetContentRows() const { return m_contentRows; }
    const int GetMatchColumnIndex() const { return m_matchColumnIndex; }
    const int GetExtractColumnIndex() const { return m_returnColumnIndex; }
    const wxArrayInt& GetExtractColumnIndices() const { return m_returnColumnIndices; }
    bool Match(const wxBaseArray<wxArrayString>& contentRows, int matchColumnIndex, const wxArrayInt& extractColumnIndices, const wxArrayString& extractColumnNames, bool enableFilter, int filterLeftColumnIndex, int filterRightColumnIndex);

private:
    void OnSelectFile(wxCommandEvent& event);
    void OnSheetChanged(wxCommandEvent& event);
    void OnAddExtractColumn(wxCommandEvent& event);
    void OnRemoveExtractColumn(wxCommandEvent& event);
    void AddExtractColumnConfigRow(const wxString& extractColumnName = wxString());
    void RefreshColumnChoices();
    bool IsFilterEnabled() const;

    wxButton* m_btnSelectFile = nullptr;
    wxTextCtrl* m_txtFilePath = nullptr;
    wxChoice* m_sheetChoice = nullptr;
    wxChoice* m_columnInputMatch = nullptr;
    wxBoxSizer* m_returnColumnsSizer = nullptr;
    wxButton* m_btnAddExtractColumn = nullptr;
    wxStaticBoxSizer* m_filterBoxSizer = nullptr;
    wxCheckBox* m_filterEnableCheck = nullptr;
    wxChoice* m_filterLeftColumnChoice = nullptr;
    wxChoice* m_filterRightColumnChoice = nullptr;

    struct ReturnColumnConfigRow
    {
        wxBoxSizer* rowSizer = nullptr;
        wxStaticText* label = nullptr;
        wxChoice* returnColumnInput = nullptr;
        wxButton* removeButton = nullptr;
    };
    std::vector<ReturnColumnConfigRow> m_returnColumnConfigRows;
    wxArrayString m_detectedColumnNames;
    wxArrayString m_titleNames;
    wxArrayString m_filterBaseColumns;

    wxString m_labelPrefix;
    bool m_loaded = false;
    xlnt::workbook m_workbook;

	bool m_bHasReturnCol = false;
    int m_matchColumnIndex = -1;
	int m_returnColumnIndex = -1;
    wxArrayInt m_returnColumnIndices;
    int m_headerRowIndex = -1;
    int m_filterLeftColumnIndex = -1;
    int m_filterRightColumnIndex = -1;
    wxBaseArray<wxArrayString> m_contentRows;
    wxArrayInt m_contentRowNumbers;

    wxString m_errorMsg;
    wxString m_outputFilePath;
};
