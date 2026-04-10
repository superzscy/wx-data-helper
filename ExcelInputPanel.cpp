#include "ExcelInputPanel.h"
#include "Logger.h"
#include <algorithm>
#include <filesystem>
#include <fstream>
#include <wx/busyinfo.h>
#include <wx/filedlg.h>

namespace
{
std::string ToUtf8String(const wxString& text)
{
    wxScopedCharBuffer utf8 = text.ToUTF8();
    if (!utf8)
    {
        return std::string();
    }
    return std::string(utf8.data());
}
}

ExcelInputPanel::ExcelInputPanel(wxWindow* parent, const wxString& labelPrefix, bool bHasReturnCol)
    : wxPanel(parent)
	, m_labelPrefix(labelPrefix)
	, m_bHasReturnCol(bHasReturnCol)
{
    wxStaticBox* box = new wxStaticBox(this, wxID_ANY, labelPrefix);
    wxStaticBoxSizer* sizer = new wxStaticBoxSizer(box, wxVERTICAL);

    // 第一行：选择文件
    wxBoxSizer* row1 = new wxBoxSizer(wxHORIZONTAL);
    m_btnSelectFile = new wxButton(this, wxID_ANY, wxT("选择文件"));
    m_txtFilePath = new wxTextCtrl(this, wxID_ANY, "", wxDefaultPosition, wxSize(300, -1), wxTE_READONLY);
    row1->Add(m_btnSelectFile, 0, wxRIGHT, 5);
    row1->Add(m_txtFilePath, 1, wxEXPAND);
    sizer->Add(row1, 0, wxEXPAND | wxBOTTOM, 5);

    // 第二行：工作表选择
    wxBoxSizer* row2 = new wxBoxSizer(wxHORIZONTAL);
    wxStaticText* sheetLabel = new wxStaticText(this, wxID_ANY, wxT("工作表："));
    m_sheetChoice = new wxChoice(this, wxID_ANY);
    row2->Add(sheetLabel, 0, wxALIGN_CENTER_VERTICAL | wxRIGHT, 5);
    row2->Add(m_sheetChoice, 1, wxEXPAND);
    sizer->Add(row2, 0, wxEXPAND | wxBOTTOM, 5);

    // 第三行：匹配列
    wxBoxSizer* row3 = new wxBoxSizer(wxHORIZONTAL);
    wxStaticText* colLabel3 = new wxStaticText(this, wxID_ANY, wxT("匹配列："));
    m_columnInputMatch = new wxChoice(this, wxID_ANY);
    row3->Add(colLabel3, 0, wxALIGN_CENTER_VERTICAL | wxRIGHT, 5);
    row3->Add(m_columnInputMatch, 1, wxEXPAND);
    sizer->Add(row3, 0, wxEXPAND | wxBOTTOM, 5);

    if (m_bHasReturnCol)
    {
        wxStaticText* cfgLabel = new wxStaticText(this, wxID_ANY, wxT("提取列："));
        sizer->Add(cfgLabel, 0, wxBOTTOM, 5);

        m_returnColumnsSizer = new wxBoxSizer(wxVERTICAL);
        sizer->Add(m_returnColumnsSizer, 0, wxEXPAND | wxBOTTOM, 5);
        AddExtractColumnConfigRow();

        wxBoxSizer* btnRow = new wxBoxSizer(wxHORIZONTAL);
        m_btnAddExtractColumn = new wxButton(this, wxID_ANY, wxT("添加提取列"));
        btnRow->Add(m_btnAddExtractColumn, 0, wxRIGHT, 5);
        sizer->Add(btnRow, 0, wxBOTTOM, 5);

        wxStaticBox* filterBox = new wxStaticBox(this, wxID_ANY, wxT("过滤数据"));
        m_filterBoxSizer = new wxStaticBoxSizer(filterBox, wxVERTICAL);

        m_filterEnableCheck = new wxCheckBox(this, wxID_ANY, wxT("启用过滤（两列值不相等时清空新加列）"));
        m_filterBoxSizer->Add(m_filterEnableCheck, 0, wxBOTTOM, 5);

        wxBoxSizer* filterRow = new wxBoxSizer(wxHORIZONTAL);
        wxStaticText* leftLabel = new wxStaticText(this, wxID_ANY, wxT("列1："));
        wxStaticText* rightLabel = new wxStaticText(this, wxID_ANY, wxT("列2："));
        m_filterLeftColumnChoice = new wxChoice(this, wxID_ANY);
        m_filterRightColumnChoice = new wxChoice(this, wxID_ANY);
        filterRow->Add(leftLabel, 0, wxALIGN_CENTER_VERTICAL | wxRIGHT, 5);
        filterRow->Add(m_filterLeftColumnChoice, 1, wxRIGHT, 10);
        filterRow->Add(rightLabel, 0, wxALIGN_CENTER_VERTICAL | wxRIGHT, 5);
        filterRow->Add(m_filterRightColumnChoice, 1);
        m_filterBoxSizer->Add(filterRow, 0, wxEXPAND);

        sizer->Add(m_filterBoxSizer, 0, wxEXPAND | wxBOTTOM, 5);
    }

    this->SetSizer(sizer);

    m_btnSelectFile->Bind(wxEVT_BUTTON, &ExcelInputPanel::OnSelectFile, this);
    if (m_btnAddExtractColumn)
    {
        m_btnAddExtractColumn->Bind(wxEVT_BUTTON, &ExcelInputPanel::OnAddExtractColumn, this);
    }
    m_sheetChoice->Bind(wxEVT_CHOICE, &ExcelInputPanel::OnSheetChanged, this);
}

void ExcelInputPanel::AddExtractColumnConfigRow(const wxString& extractColumnName)
{
    if (!m_returnColumnsSizer)
    {
        return;
    }

    wxBoxSizer* rowSizer = new wxBoxSizer(wxHORIZONTAL);
    wxStaticText* rowLabel = new wxStaticText(this, wxID_ANY, wxT("提取列："));
    wxChoice* returnColInput = new wxChoice(this, wxID_ANY);
    wxButton* removeBtn = new wxButton(this, wxID_ANY, wxT("删除"));
    rowSizer->Add(rowLabel, 0, wxALIGN_CENTER_VERTICAL | wxRIGHT, 5);
    rowSizer->Add(returnColInput, 1, wxRIGHT, 8);
    rowSizer->Add(removeBtn, 0);

    m_returnColumnsSizer->Add(rowSizer, 0, wxEXPAND | wxBOTTOM, 4);
    m_returnColumnConfigRows.push_back({ rowSizer, rowLabel, returnColInput, removeBtn });
    removeBtn->Bind(wxEVT_BUTTON, &ExcelInputPanel::OnRemoveExtractColumn, this);

    for (const auto& colName : m_detectedColumnNames)
    {
        returnColInput->Append(colName);
    }
    if (!extractColumnName.IsEmpty())
    {
        int idx = returnColInput->FindString(extractColumnName);
        if (idx != wxNOT_FOUND)
        {
            returnColInput->SetSelection(idx);
        }
    }
    else if (returnColInput->GetCount() > 0)
    {
        returnColInput->SetSelection(0);
    }
}

void ExcelInputPanel::OnAddExtractColumn(wxCommandEvent& WXUNUSED(event))
{
    AddExtractColumnConfigRow();
    Layout();
    GetParent()->Layout();

    wxTopLevelWindow* top = wxDynamicCast(wxGetTopLevelParent(this), wxTopLevelWindow);
    if (top)
    {
        wxSize size = top->GetSize();
        top->SetSize(size.GetWidth(), size.GetHeight() + 42);
    }
}

void ExcelInputPanel::OnRemoveExtractColumn(wxCommandEvent& event)
{
    wxObject* src = event.GetEventObject();
    for (size_t i = 0; i < m_returnColumnConfigRows.size(); i++)
    {
        auto& row = m_returnColumnConfigRows[i];
        if (row.removeButton == src)
        {
            if (row.rowSizer)
            {
                m_returnColumnsSizer->Detach(row.rowSizer);
                delete row.rowSizer;
            }
            if (row.label)
            {
                row.label->Destroy();
            }
            if (row.returnColumnInput)
            {
                row.returnColumnInput->Destroy();
            }
            if (row.removeButton)
            {
                row.removeButton->Destroy();
            }
            m_returnColumnConfigRows.erase(m_returnColumnConfigRows.begin() + i);
            Layout();
            GetParent()->Layout();

            wxTopLevelWindow* top = wxDynamicCast(wxGetTopLevelParent(this), wxTopLevelWindow);
            if (top)
            {
                const int deltaHeight = 42;
                const int currentHeight = top->GetSize().GetHeight();
                int minHeight = top->GetBestSize().GetHeight();
                if (minHeight <= 0)
                {
                    minHeight = top->GetMinSize().GetHeight();
                }
                if (minHeight <= 0)
                {
                    minHeight = 320;
                }
                const int newHeight = std::max(minHeight, currentHeight - deltaHeight);
                if (newHeight != currentHeight)
                {
                    top->SetSize(top->GetSize().GetWidth(), newHeight);
                }
            }
            break;
        }
    }
}

void ExcelInputPanel::OnSheetChanged(wxCommandEvent& WXUNUSED(event))
{
    RefreshColumnChoices();
}

void ExcelInputPanel::RefreshColumnChoices()
{
    m_detectedColumnNames.Clear();
    if (!m_loaded || GetSelectedSheet().IsEmpty())
    {
        return;
    }

    try
    {
        auto ws = m_workbook.sheet_by_title(ToUtf8String(GetSelectedSheet()));
        int highestColumn = ws.highest_column().index;
        int highestRow = ws.highest_row();

        for (int row = 1; row <= highestRow; row++)
        {
            wxArrayString candidates;
            bool hasValue = false;
            for (int col = 1; col <= highestColumn; col++)
            {
                xlnt::cell cell = ws.cell(xlnt::cell_reference(col, row));
                wxString content;
                if (cell.has_value())
                {
                    content = wxString::FromUTF8(cell.to_string());
                }
                content = content.Trim(true).Trim(false);
                candidates.Add(content);
                if (!content.IsEmpty())
                {
                    hasValue = true;
                }
            }
            if (hasValue)
            {
                m_detectedColumnNames = candidates;
                break;
            }
        }
    }
    catch (...)
    {
    }

    wxString currentMatch = GetMatchColumnName();
    m_columnInputMatch->Clear();
    for (const auto& colName : m_detectedColumnNames)
    {
        m_columnInputMatch->Append(colName);
    }
    if (!currentMatch.IsEmpty())
    {
        int idx = m_columnInputMatch->FindString(currentMatch);
        if (idx != wxNOT_FOUND)
        {
            m_columnInputMatch->SetSelection(idx);
        }
    }
    if (m_columnInputMatch->GetSelection() == wxNOT_FOUND && m_columnInputMatch->GetCount() > 0)
    {
        m_columnInputMatch->SetSelection(0);
    }

    for (auto& row : m_returnColumnConfigRows)
    {
        if (!row.returnColumnInput)
        {
            continue;
        }
        wxString current = row.returnColumnInput->GetStringSelection();
        row.returnColumnInput->Clear();
        for (const auto& colName : m_detectedColumnNames)
        {
            row.returnColumnInput->Append(colName);
        }
        if (!current.IsEmpty())
        {
            int idx = row.returnColumnInput->FindString(current);
            if (idx != wxNOT_FOUND)
            {
                row.returnColumnInput->SetSelection(idx);
            }
        }
        if (row.returnColumnInput->GetSelection() == wxNOT_FOUND && row.returnColumnInput->GetCount() > 0)
        {
            row.returnColumnInput->SetSelection(0);
        }
    }

    if (m_filterLeftColumnChoice && m_filterRightColumnChoice)
    {
        wxString leftCurrent = m_filterLeftColumnChoice->GetStringSelection();
        wxString rightCurrent = m_filterRightColumnChoice->GetStringSelection();

        m_filterLeftColumnChoice->Clear();
        m_filterRightColumnChoice->Clear();

        wxArrayString filterColumns = (m_filterBaseColumns.GetCount() == 0) ? m_detectedColumnNames : m_filterBaseColumns;
        if (m_bHasReturnCol)
        {
            wxArrayString extractNames = GetExtractColumnNames();
            for (const auto& colName : extractNames)
            {
                if (filterColumns.Index(colName, false) == wxNOT_FOUND)
                {
                    filterColumns.Add(colName);
                }
            }
        }

        for (const auto& colName : filterColumns)
        {
            m_filterLeftColumnChoice->Append(colName);
            m_filterRightColumnChoice->Append(colName);
        }

        if (!leftCurrent.IsEmpty())
        {
            int idx = m_filterLeftColumnChoice->FindString(leftCurrent);
            if (idx != wxNOT_FOUND)
            {
                m_filterLeftColumnChoice->SetSelection(idx);
            }
        }
        if (!rightCurrent.IsEmpty())
        {
            int idx = m_filterRightColumnChoice->FindString(rightCurrent);
            if (idx != wxNOT_FOUND)
            {
                m_filterRightColumnChoice->SetSelection(idx);
            }
        }
    }
}

void ExcelInputPanel::ConfigureFilterBaseColumns(const wxArrayString& baseColumns)
{
    m_filterBaseColumns = baseColumns;
    RefreshColumnChoices();
}

void ExcelInputPanel::OnSelectFile(wxCommandEvent& WXUNUSED(event))
{
    wxFileDialog dlg(this, wxT("选择工作簿文件"), "", "", wxT("工作簿文件 (*.xlsx)|*.xlsx"), wxFD_OPEN | wxFD_FILE_MUST_EXIST);
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

        RefreshColumnChoices();
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
    return m_columnInputMatch ? m_columnInputMatch->GetStringSelection() : wxString("");
}

wxString ExcelInputPanel::GetExtractColumnName() const {
    wxArrayString names = GetExtractColumnNames();
    return names.GetCount() == 0 ? wxString("") : names[0];
}

wxArrayString ExcelInputPanel::GetExtractColumnNames() const
{
    wxArrayString names;
    for (const auto& row : m_returnColumnConfigRows)
    {
        if (!row.returnColumnInput)
        {
            continue;
        }

        wxString returnName = row.returnColumnInput->GetStringSelection().Trim(true).Trim(false);
        if (!returnName.IsEmpty())
        {
            names.Add(returnName);
        }
    }
    return names;
}

bool ExcelInputPanel::IsFilterEnabled() const
{
    if (!m_filterEnableCheck || !m_filterEnableCheck->IsChecked())
    {
        return false;
    }

    if (!m_filterLeftColumnChoice || !m_filterRightColumnChoice)
    {
        return false;
    }

    return !m_filterLeftColumnChoice->GetStringSelection().IsEmpty() &&
        !m_filterRightColumnChoice->GetStringSelection().IsEmpty();
}

bool ExcelInputPanel::Match(const wxBaseArray<wxArrayString>& contentRows, int matchColumnIndex, const wxArrayInt& extractColumnIndices, const wxArrayString& extractColumnNames, bool enableFilter, int filterLeftColumnIndex, int filterRightColumnIndex)
{
    if (m_contentRows.size() != m_contentRowNumbers.GetCount())
    {
        m_errorMsg = wxT("内部行号状态异常");
        return false;
    }
    if (extractColumnIndices.GetCount() == 0 || extractColumnIndices.GetCount() != extractColumnNames.GetCount())
    {
        m_errorMsg = wxT("提取列配置无效");
        return false;
    }

    std::vector<wxArrayString> matchedColumns;
    matchedColumns.resize(extractColumnIndices.GetCount());
    LOG_INFO(wxString::Format(wxT("开始提取，共%d行"), (int)m_contentRows.size()).ToUTF8().data());
    const int baseColumnCount = (m_filterBaseColumns.GetCount() == 0) ? m_titleNames.GetCount() : m_filterBaseColumns.GetCount();

    for (int index = 0; index < m_contentRows.size(); index++)
    {
        const auto& myRow = m_contentRows[index];
        wxString myKey = myRow[m_matchColumnIndex];
        const wxArrayString* foundOtherRow = nullptr;

        for (const auto& theOtherRow : contentRows)
        {
            wxString theOtherKey = theOtherRow[matchColumnIndex];
            if(myKey == theOtherKey)
            {
                foundOtherRow = &theOtherRow;
                break;
			}
        }
        wxArrayString currentExtractValues;

        for (int colIdx = 0; colIdx < extractColumnIndices.GetCount(); colIdx++)
        {
            wxString searchValue;
            const int extractColumnIndex = extractColumnIndices[colIdx];
            if (foundOtherRow && extractColumnIndex >= 0 && extractColumnIndex < (int)foundOtherRow->GetCount())
            {
                searchValue = (*foundOtherRow)[extractColumnIndex];
            }
            matchedColumns[colIdx].Add(searchValue);
            currentExtractValues.Add(searchValue);
        }

        if (enableFilter)
        {
            auto getFilterValue = [&](int combinedIndex) -> wxString
            {
                if (combinedIndex < 0)
                {
                    return wxString("");
                }
                if (combinedIndex < baseColumnCount)
                {
                    if (combinedIndex < (int)myRow.GetCount())
                    {
                        return myRow[combinedIndex];
                    }
                    return wxString("");
                }
                int extractIdx = combinedIndex - baseColumnCount;
                if (extractIdx >= 0 && extractIdx < (int)currentExtractValues.GetCount())
                {
                    return currentExtractValues[extractIdx];
                }
                return wxString("");
            };

            wxString leftValue = getFilterValue(filterLeftColumnIndex);
            wxString rightValue = getFilterValue(filterRightColumnIndex);
            if (leftValue != rightValue)
            {
                for (int colIdx = 0; colIdx < (int)matchedColumns.size(); colIdx++)
                {
                    matchedColumns[colIdx][index].clear();
                }
            }
        }
    }

    if (enableFilter)
    {
        LOG_INFO("过滤完成：不满足条件的行已清空新加列");
    }

    LOG_INFO("提取完成，结果：");
    wxString fullText;
    for (int colIdx = 0; colIdx < (int)matchedColumns.size(); colIdx++)
    {
        fullText += wxString::Format(wxT("提取列%d(%s): "), colIdx + 1, extractColumnNames[colIdx]);
        for (const auto& val : matchedColumns[colIdx])
        {
            fullText += val + ",";
        }
        fullText += "\n";
    }
    LOG_INFO(wxString::Format(wxT("\n---\n%s---"), fullText).ToUTF8().data());

    try
    {
        const std::filesystem::path sourcePath(GetFilePath().ToStdWstring());
        std::filesystem::path outputPath = sourcePath.parent_path() /
            std::filesystem::path(sourcePath.stem().wstring() + L"_matched" + sourcePath.extension().wstring());

        std::filesystem::copy_file(sourcePath, outputPath, std::filesystem::copy_options::overwrite_existing);

        xlnt::workbook outputWorkbook;
        {
            std::ifstream file(outputPath, std::ios::binary);
            outputWorkbook.load(file);
        }

        auto ws = outputWorkbook.sheet_by_title(ToUtf8String(GetSelectedSheet()));
        int targetStartCol = ws.highest_column().index + 1;

        for (int colIdx = 0; colIdx < extractColumnIndices.GetCount(); colIdx++)
        {
            int targetCol = targetStartCol + colIdx;
            if (m_headerRowIndex > 0)
            {
                ws.cell(xlnt::cell_reference(targetCol, m_headerRowIndex)).value(ToUtf8String(extractColumnNames[colIdx]));
            }

            for (int i = 0; i < matchedColumns[colIdx].GetCount(); i++)
            {
                ws.cell(xlnt::cell_reference(targetCol, m_contentRowNumbers[i])).value(ToUtf8String(matchedColumns[colIdx][i]));
            }
        }

        {
            std::ofstream outFile(outputPath, std::ios::binary);
            outputWorkbook.save(outFile);
        }

        m_outputFilePath = wxString(outputPath.wstring().c_str());
        LOG_INFO(wxString::Format(wxT("结果已写入：%s"), m_outputFilePath).ToUTF8().data());
        return true;
    }
    catch (const std::exception& e)
    {
        m_errorMsg = wxString::Format(wxT("写入结果失败：%s"), e.what());
        return false;
    }
}

bool ExcelInputPanel::IsReady() const {
    if (!m_loaded || GetSelectedSheet().IsEmpty() || GetMatchColumnName().IsEmpty())
    {
        return false;
    }

    if (!m_bHasReturnCol)
    {
        return true;
    }

    bool hasValidConfig = false;
    for (const auto& row : m_returnColumnConfigRows)
    {
        wxString returnName = row.returnColumnInput ? row.returnColumnInput->GetStringSelection().Trim(true).Trim(false) : wxString();
        if (!returnName.IsEmpty())
        {
            hasValidConfig = true;
        }
    }
    return hasValidConfig;
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
        m_contentRowNumbers.Empty();
        m_titleNames.Clear();
        m_matchColumnIndex = -1;
        m_returnColumnIndex = -1;
        m_returnColumnIndices.Empty();
        m_headerRowIndex = -1;
        m_filterLeftColumnIndex = -1;
        m_filterRightColumnIndex = -1;
        m_outputFilePath.clear();

        auto ws = m_workbook.sheet_by_title(ToUtf8String(GetSelectedSheet()));

        wxArrayString titleNames;
        wxArrayString returnColumnNames = GetExtractColumnNames();

        if (m_bHasReturnCol)
        {
            if (returnColumnNames.GetCount() == 0)
            {
                m_errorMsg = wxString::Format(wxT("%s中没有有效的提取列"), m_labelPrefix);
                return false;
            }
            m_returnColumnIndices.Alloc(returnColumnNames.GetCount());
            for (int i = 0; i < returnColumnNames.GetCount(); i++)
            {
                m_returnColumnIndices.Add(-1);
            }
        }

		int startRowIndex = 0;
        int highestColumn = ws.highest_column().index;
		int rowCount = ws.highest_row();

        auto allRows = ws.rows(true);
        LOG_INFO(wxString::Format(wxT("总行数：%d"), (int)allRows.length()).ToUTF8().data());

        for (int rowIndex = 0; rowIndex < allRows.length(); rowIndex++)
        {
            const auto& row = allRows[rowIndex];
            if (row.empty())
            {
                break;
            }
            parsedRowIndex = row.front().row();  // 当前行号

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
                    m_matchColumnIndex = col - 1; // 从0开始
                }

                if (m_bHasReturnCol)
                {
                    for (int returnIdx = 0; returnIdx < returnColumnNames.GetCount(); returnIdx++)
                    {
                        if (m_returnColumnIndices[returnIdx] == -1 && cellContent == returnColumnNames[returnIdx])
                        {
                            m_returnColumnIndices[returnIdx] = col - 1; // 从0开始
                        }
                    }
                }
            }

            LOG_INFO(wxString::Format(wxT("  行:%04d %s"), parsedRowIndex, rowContent).ToUTF8().data());

            if (m_matchColumnIndex >= 0)
            {
                if (titleNames.GetCount() == 0)
                {
                    titleNames = currentRowCntent;
                    m_titleNames = currentRowCntent;
                    m_headerRowIndex = parsedRowIndex;
                }
                else
                {
                    m_contentRows.Add(currentRowCntent);
                    m_contentRowNumbers.Add(parsedRowIndex);
                }
            }
		}

        if (m_matchColumnIndex == -1)
        {
            m_errorMsg = wxString::Format(wxT("%s中未找到匹配列"), m_labelPrefix);
            return false;
        }

        if (m_bHasReturnCol && IsFilterEnabled() && titleNames.GetCount() > 0)
        {
            wxString leftName = m_filterLeftColumnChoice->GetStringSelection();
            wxString rightName = m_filterRightColumnChoice->GetStringSelection();
            wxArrayString filterColumns = (m_filterBaseColumns.GetCount() == 0) ? titleNames : m_filterBaseColumns;
            wxArrayString extractNames = GetExtractColumnNames();
            for (const auto& colName : extractNames)
            {
                if (filterColumns.Index(colName, false) == wxNOT_FOUND)
                {
                    filterColumns.Add(colName);
                }
            }

            for (int i = 0; i < filterColumns.GetCount(); i++)
            {
                if (m_filterLeftColumnIndex == -1 && filterColumns[i] == leftName)
                {
                    m_filterLeftColumnIndex = i;
                }
                if (m_filterRightColumnIndex == -1 && filterColumns[i] == rightName)
                {
                    m_filterRightColumnIndex = i;
                }
            }

            if (m_filterLeftColumnIndex == -1 || m_filterRightColumnIndex == -1)
            {
                m_errorMsg = wxString::Format(wxT("%s中未找到过滤列"), m_labelPrefix);
                return false;
            }
        }

        if (m_bHasReturnCol && m_returnColumnIndices.GetCount() > 0)
        {
            for (int i = 0; i < m_returnColumnIndices.GetCount(); i++)
            {
                if (m_returnColumnIndices[i] == -1)
                {
                    m_errorMsg = wxString::Format(wxT("%s中未找到提取列：%s"), m_labelPrefix, returnColumnNames[i]);
                    return false;
                }
            }
            m_returnColumnIndex = m_returnColumnIndices[0];
        }

        if(m_contentRows.Count() == 0)
        {
            m_errorMsg = wxString::Format(wxT("%s中没有有效数据行"), m_labelPrefix);
            return false;
		}

        //wxLogMessage(wxT("数据起始行：%d"), startRowIndex);
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
