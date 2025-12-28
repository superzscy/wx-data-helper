#include "ExcelInputPanel.h"
#include "Logger.h"
#include <filesystem>
#include <fstream>
#include <wx/busyinfo.h>
#include <wx/clipbrd.h>
#include <wx/filedlg.h>
#include <wx/dataobj.h>

ExcelInputPanel::ExcelInputPanel(wxWindow* parent, const wxString& labelPrefix, bool bHasReturnCol)
    : wxPanel(parent)
	, m_labelPrefix(labelPrefix)
	, m_bHasReturnCol(bHasReturnCol)
{
    wxStaticBox* box = new wxStaticBox(this, wxID_ANY, labelPrefix);
    wxStaticBoxSizer* sizer = new wxStaticBoxSizer(box, wxVERTICAL);

    // 第一行：按钮 + 文件路径
    wxBoxSizer* row1 = new wxBoxSizer(wxHORIZONTAL);
    m_btnSelectFile = new wxButton(this, wxID_ANY, wxT("选择文件"));
    m_txtFilePath = new wxTextCtrl(this, wxID_ANY, "", wxDefaultPosition, wxSize(300, -1), wxTE_READONLY);
    row1->Add(m_btnSelectFile, 0, wxRIGHT, 5);
    row1->Add(m_txtFilePath, 1, wxEXPAND);
    sizer->Add(row1, 0, wxEXPAND | wxBOTTOM, 5);

    // 第二行：下拉 sheet 选择
    wxBoxSizer* row2 = new wxBoxSizer(wxHORIZONTAL);
    wxStaticText* sheetLabel = new wxStaticText(this, wxID_ANY, wxT("选择工作表："));
    m_sheetChoice = new wxChoice(this, wxID_ANY);
    row2->Add(sheetLabel, 0, wxALIGN_CENTER_VERTICAL | wxRIGHT, 5);
    row2->Add(m_sheetChoice, 1, wxEXPAND);
    sizer->Add(row2, 0, wxEXPAND | wxBOTTOM, 5);

    // 第三行：输入列名
    wxBoxSizer* row3 = new wxBoxSizer(wxHORIZONTAL);
    wxStaticText* colLabel3 = new wxStaticText(this, wxID_ANY, wxT("匹配列名："));
    m_columnInputMatch = new wxTextCtrl(this, wxID_ANY, "", wxDefaultPosition, wxDefaultSize);
    row3->Add(colLabel3, 0, wxALIGN_CENTER_VERTICAL | wxRIGHT, 5);
    row3->Add(m_columnInputMatch, 1, wxEXPAND);
    sizer->Add(row3, 0, wxEXPAND | wxBOTTOM, 5);

    if (m_bHasReturnCol)
    {
        // 第四行：输入列名
        wxBoxSizer* row4 = new wxBoxSizer(wxHORIZONTAL);
        wxStaticText* colLabel4 = new wxStaticText(this, wxID_ANY, wxT("返回列名："));
        m_columnInputReturn = new wxTextCtrl(this, wxID_ANY, "", wxDefaultPosition, wxDefaultSize);
        row4->Add(colLabel4, 0, wxALIGN_CENTER_VERTICAL | wxRIGHT, 5);
        row4->Add(m_columnInputReturn, 1, wxEXPAND);
        sizer->Add(row4, 0, wxEXPAND | wxBOTTOM, 5);
    }

    this->SetSizer(sizer);

    m_btnSelectFile->Bind(wxEVT_BUTTON, &ExcelInputPanel::OnSelectFile, this);
}

void ExcelInputPanel::OnSelectFile(wxCommandEvent& WXUNUSED(event))
{
    wxFileDialog dlg(this, wxT("选择 Excel 文件"), "", "", wxT("Excel 文件 (*.xlsx)|*.xlsx"), wxFD_OPEN | wxFD_FILE_MUST_EXIST);
    if (dlg.ShowModal() == wxID_CANCEL)
        return;

    wxString filepath = dlg.GetPath();
    m_txtFilePath->SetValue(filepath);

    try {
        wxBusyInfo info(wxT("正在读取文件，请稍候..."), GetParent());

        std::ifstream file(std::filesystem::path(filepath.ToStdWstring()), std::ios::binary);
        m_workbook.load(file);
        m_loaded = true;

        m_sheetChoice->Clear();
        for (const auto& name : m_workbook.sheet_titles()) {
            m_sheetChoice->Append(wxString::FromUTF8(name));
        }

        if (m_sheetChoice->GetCount() > 0)
            m_sheetChoice->SetSelection(0);
    }
    catch (const std::exception& e) {
        wxMessageBox(wxString::Format(wxT("打开失败：%s"), e.what()), wxT("错误"), wxICON_ERROR);
        m_loaded = false;
    }
}

wxString ExcelInputPanel::GetFilePath() const {
    return m_txtFilePath->GetValue();
}

wxString ExcelInputPanel::GetSelectedSheet() const {
    return m_sheetChoice->GetStringSelection();
}

wxString ExcelInputPanel::GetMatchColumnName() const {
    return m_columnInputMatch->GetValue();
}

wxString ExcelInputPanel::GetReturnColumnName() const {
    return m_columnInputReturn ? m_columnInputReturn->GetValue() : wxString("");
}

bool ExcelInputPanel::Match(const wxBaseArray<wxArrayString>& contentRows, int matchColumnIndex, int returnColumnIndex)
{
    wxString fullText;
    LOG_INFO(wxString::Format(wxT("开始查找, 一共%d行"), (int)m_contentRows.size()).ToUTF8().data());

    for (int index = 0; index < m_contentRows.size(); index++)
    {
        const auto& myRow = m_contentRows[index];

        wxString searchValue;
        wxString myKey = myRow[m_matchColumnIndex];

        for (const auto& theOtherRow : contentRows)
        {
            wxString theOtherKey = theOtherRow[matchColumnIndex];
            if(myKey == theOtherKey)
            {
                searchValue = theOtherRow[returnColumnIndex];
                break;
			}
        }
        fullText += searchValue + "\n";
    }

    LOG_INFO("查找完成, 结果:");
    LOG_INFO(wxString::Format(wxT("\n---\n%s---"), fullText).ToUTF8().data());
    
    for (int i = 0; i < 3; i++)
    {
        wxTheClipboard->Clear();  // 清除旧内容
        wxTheClipboard->SetData(new wxTextDataObject(fullText));
        if (wxTheClipboard->Flush())
        {
            return true;
        }

        wxThread::Sleep(1000);
    }

    m_errorMsg = wxString(wxT("无法将数据复制到剪贴板, 请从日志里手动拷贝结果"));
    return false;
}

bool ExcelInputPanel::IsReady() const {
    return m_loaded && !GetSelectedSheet().IsEmpty() && !GetMatchColumnName().IsEmpty() && (!m_bHasReturnCol or !GetReturnColumnName().IsEmpty());
}

bool ExcelInputPanel::Parse() {
    if (!m_loaded)
    {
        m_errorMsg = wxString::Format(wxT("%s未加载"), m_labelPrefix);
        return false;
    }
    LOG_INFO(wxString::Format(wxT("解析%s"), GetFilePath()).ToUTF8().data());

    int parsedRowIndex = 0;
    try
    {
        m_errorMsg.Clear();
		m_contentRows.Empty();
        m_matchColumnIndex = -1;
        m_returnColumnIndex = -1;

        auto ws = m_workbook.sheet_by_title(GetSelectedSheet().ToStdString());

        wxArrayString titleNames;

		int startRowIndex = 0;
        int highestColumn = ws.highest_column().index;
		int rowCount = ws.highest_row();

        auto allRows = ws.rows(true);
        LOG_INFO(wxString::Format(wxT("总共行数:%d"), (int)allRows.length()).ToUTF8().data());

        for (int rowIndex = 0; rowIndex < allRows.length(); rowIndex++)
        {
            const auto& row = allRows[rowIndex];
            if (row.empty())
            {
                break;
            }
            parsedRowIndex = row.front().row();  // 获取当前行号

            wxArrayString currentRowCntent;
            wxString rowContent;
            for (int col = 1; col <= highestColumn; ++col)
            {
                xlnt::cell cell = ws.cell(xlnt::cell_reference(col, parsedRowIndex));
				wxString cellContent;
                if (cell.has_value())
                {
                    cellContent = wxString::FromUTF8(cell.to_string());
                }

                rowContent += " | " + cellContent;

                currentRowCntent.Add(cellContent);
                if ((m_matchColumnIndex == -1) && cellContent == GetMatchColumnName())
                {
					startRowIndex = parsedRowIndex;
                    m_matchColumnIndex = col - 1; // 索引从0开始
                }

                if (m_bHasReturnCol && (m_returnColumnIndex == -1) && cellContent == GetReturnColumnName())
                {
                    m_returnColumnIndex = col - 1; // 索引从0开始
                }
            }

            LOG_INFO(wxString::Format(wxT("  Line:%04d %s"), parsedRowIndex, rowContent).ToUTF8().data());

            if (m_matchColumnIndex >= 0)
            {
                if (titleNames.GetCount() == 0)
                {
                    titleNames = currentRowCntent;
                }
                else
                {
                    m_contentRows.Add(currentRowCntent);
                }
            }
		}

        if (m_matchColumnIndex == -1)
        {
            m_errorMsg = wxString::Format(wxT("%s中没有找到匹配列名"), m_labelPrefix);
            return false;
        }

        if (m_bHasReturnCol && m_returnColumnIndex == -1)
        {
            m_errorMsg = wxString::Format(wxT("%s中没有找到返回列名"), m_labelPrefix);
            return false;
        }

        if(m_contentRows.Count() == 0)
        {
            m_errorMsg = wxString::Format(wxT("%s中没有找到有效数据"), m_labelPrefix);
            return false;
		}

        //wxLogMessage(wxT("数据起始行: %d"), startRowIndex);
        wxString titleContent;
        for (const auto& col : titleNames)
        {
            titleContent += col + ";";
        }
        //wxLogMessage(titleContent);

        for (const auto& row : m_contentRows)
        {
            wxString content;
            for (const auto& col : row)
            {
                content += col + ";";
            }
            //wxLogMessage(content);
        }

		return true;
    }
    catch (const std::exception& e) {
        m_errorMsg = wxString::Format(wxT("%s解析失败：%d %s"), m_labelPrefix, parsedRowIndex, e.what());
        return false;
	}
}